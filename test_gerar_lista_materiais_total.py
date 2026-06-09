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


class MergeSheetTests(unittest.TestCase):
    def test_build_merge_sheet_writes_elem_values_under_elemento_header(self):
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

        ws = lm._build_merge_sheet(wb, "SEG 1", data, [8.0])

        self.assertEqual("ELEMENTO", ws.cell(row=1, column=1).value)
        self.assertEqual("P1", ws.cell(row=2, column=1).value)


if __name__ == "__main__":
    unittest.main()
