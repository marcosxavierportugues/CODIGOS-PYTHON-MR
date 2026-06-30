import os
import tempfile
import unittest
from pathlib import Path

import pandas as pd

import gerar_lista_materiais_total as lm


class LoadQtoTests(unittest.TestCase):
    def test_load_qto_reads_raw_ememento_column_before_material_parsing(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "qto.xlsx"
            pd.DataFrame(
                {
                    "Layer Name": ["NAO_E_ELEMENTO"],
                    "Element Id": [101],
                    "Element Type": ["PILAR"],
                    "Ememento": ["P1"],
                    "Layer": ["TERREO"],
                    "Descricao 1": ["Forma - Estrutura"],
                    "Quantidade 1": ["12,50"],
                    "Descricao 2": ["Armadura CA50 8.0 mm"],
                    "Quantidade 2": ["7,20"],
                }
            ).to_excel(path, index=False)

            data, bitolas, has_layer_col = lm.load_qto(path, default_nivel="PAV")

        self.assertTrue(has_layer_col)
        self.assertEqual(["P1"], data["ELEM"].tolist())
        self.assertEqual(["TERREO"], data["NIVEL"].tolist())
        self.assertAlmostEqual(12.5, data.loc[0, "FORMA_m2"])
        self.assertEqual([8.0], bitolas)
        self.assertAlmostEqual(7.2, data.loc[0, "ACO_8.0mm_kg"])


class SegmentosTests(unittest.TestCase):
    def test_load_segmentos_reads_ememento_instead_of_layer_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "seg1.xlsx"
            pd.DataFrame(
                {
                    "Layer Name": ["TERREO", "BALDRAME"],
                    "Element Type": ["PILAR", "PILAR"],
                    "Ememento": ["P1", "P2"],
                }
            ).to_excel(path, index=False)

            data = lm.load_segmentos({"SEG 1": path})

        self.assertEqual(["P1", "P2"], data["ELEM"].tolist())
        self.assertEqual(["SEG 1", "SEG 1"], data["SEGMENTO"].tolist())


class MergeMetadataTests(unittest.TestCase):
    def test_coletar_metadados_uniao_asks_fields_for_frontend(self):
        original_ask_str = lm._ask_str
        original_ask_choice = lm._ask_choice
        original_env = {name: os.environ.get(name) for name in ("LM_ELEMENTO", "LM_PAVIMENTO", "LM_OBRA", "LM_TIPO", "LM_DATA")}
        calls = []
        choice_calls = []
        answers = {
            "Elemento": "pilares",
            "Pavimento": "terreo",
            "Obra": "ifmt",
            "Data": "09.06.2026",
        }

        def fake_ask_str(title, prompt, default=""):
            calls.append((title, prompt, default))
            return answers[title]

        def fake_ask_choice(title, prompt, options):
            choice_calls.append((title, prompt, options))
            return "EXECUTIVO"

        try:
            for name in original_env:
                os.environ.pop(name, None)
            lm._ask_str = fake_ask_str
            lm._ask_choice = fake_ask_choice

            meta = lm.coletar_metadados_uniao()
        finally:
            lm._ask_str = original_ask_str
            lm._ask_choice = original_ask_choice
            for name, value in original_env.items():
                if value is None:
                    os.environ.pop(name, None)
                else:
                    os.environ[name] = value

        self.assertEqual(["Elemento", "Pavimento", "Obra", "Data"], [c[0] for c in calls])
        self.assertEqual(["Tipo"], [c[0] for c in choice_calls])
        self.assertEqual("PILARES", meta["elemento"])
        self.assertEqual("TERREO", meta["pavimento"])
        self.assertEqual("IFMT", meta["obra"])
        self.assertEqual("EXECUTIVO", meta["tipo"])
        self.assertEqual("09.06.2026", meta["data"])
        self.assertEqual("UNIAO LISTAS - PILARES - TERREO - IFMT - EXECUTIVO - 09.06.2026", meta["nome_base"])

    def test_build_merge_sheet_adds_metadata_title_above_table(self):
        wb = lm.Workbook()
        active = wb.active
        if active is not None:
            wb.remove(active)
        data = pd.DataFrame(
            {
                "ELEM": ["P1"],
                "NIVEL": ["TERREO"],
                "CONCRETO_m3": [1.25],
                "FORMA_m2": [2.5],
                "ACO_8.0mm_kg": [3.75],
            }
        )

        ws = lm._build_merge_sheet(wb, "TODOS - SAPATAS SEG 1", data, [8.0], self._meta())

        self.assertEqual("UNIAO DE LISTAS - SAPATAS - TERREO - IFMT - EXECUTIVO - 09.06.2026 - TODOS - SAPATAS SEG 1", ws["A1"].value)
        self.assertEqual("ELEMENTO", ws.cell(row=3, column=1).value)
        self.assertEqual("P1", ws.cell(row=4, column=1).value)
        self.assertEqual("A4", ws.freeze_panes)

    def test_build_merge_summary_and_total_add_metadata_titles_above_tables(self):
        wb = lm.Workbook()
        active = wb.active
        if active is not None:
            wb.remove(active)
        summary = {
            "segment": "SEG 1",
            "group_label": "VIGAS",
            "totals": {"CONCRETO_m3": 1.0, "FORMA_m2": 2.0, "ACO_8.0mm_kg": 3.0},
        }

        ws_summary = lm._build_merge_summary_sheet(wb, [summary], [8.0], self._meta())
        ws_total = lm._build_merge_total_sheet(
            wb,
            {"CONCRETO_m3": 1.0, "FORMA_m2": 2.0, "ACO_8.0mm_kg": 3.0},
            "SAPATAS",
            [8.0],
            self._meta(),
        )

        self.assertEqual("UNIAO DE LISTAS - SAPATAS - TERREO - IFMT - EXECUTIVO - 09.06.2026 - RESUMO POR SEGMENTO", ws_summary["A1"].value)
        self.assertEqual("SEGMENTO", ws_summary.cell(row=3, column=1).value)
        self.assertEqual("SAPATAS", ws_summary.cell(row=4, column=2).value)
        self.assertEqual("TOTAL GERAL", ws_summary.cell(row=5, column=1).value)
        self.assertEqual("SAPATAS", ws_summary.cell(row=5, column=2).value)
        self.assertEqual("UNIAO DE LISTAS - SAPATAS - TERREO - IFMT - EXECUTIVO - 09.06.2026 - RESUMO TOTAL SAPATAS", ws_total["A1"].value)
        self.assertEqual("ELEMENTO", ws_total.cell(row=3, column=1).value)
        self.assertEqual("SAPATAS", ws_total.cell(row=4, column=1).value)

    def _meta(self):
        return {
            "elemento": "SAPATAS",
            "pavimento": "TERREO",
            "obra": "IFMT",
            "tipo": "EXECUTIVO",
            "data": "09.06.2026",
            "nome_base": "UNIAO LISTAS - SAPATAS - TERREO - IFMT - EXECUTIVO - 09.06.2026",
        }


if __name__ == "__main__":
    unittest.main()
