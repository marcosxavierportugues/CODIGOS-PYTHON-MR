import importlib.machinery
import importlib.util
import tempfile
import unittest
from pathlib import Path

import pandas as pd


def _load_teste3_module():
    path = Path(__file__).with_name("TESTE 3.PY")
    loader = importlib.machinery.SourceFileLoader("teste3", str(path))
    spec = importlib.util.spec_from_loader("teste3", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class Teste3ItemColumnTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.mod = _load_teste3_module()

    def test_find_item_column_accepts_item_with_line_break_case_and_accent_variation(self):
        cols = [
            "AltoQi_Eberick_Elemento / Elemento",
            "\u00cdtEm\r\nnema",
            "Descri\u00e7\u00e3o 1",
            "Quantidade 1",
        ]

        item_col = self.mod._find_item_column(cols, allow_legacy=False)

        self.assertEqual("\u00cdtEm\r\nnema", item_col)

    def test_load_qto_uses_item_column_with_line_break_and_ignores_technical_element_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "qto.xlsx"
            pd.DataFrame(
                {
                    "AltoQi_Eberick_Elemento / Elemento": ["NAO_USAR"],
                    "\u00cdtEm\r\nnema": ["SAPATA S1"],
                    "Layer": ["BALDRAME"],
                    "Descri\u00e7\u00e3o 1": ["Forma - Estrutura"],
                    "Quantidade 1": ["12,5"],
                }
            ).to_excel(path, index=False)

            data, bitolas, has_layer_col = self.mod.load_qto(path, default_nivel="PAV")

        self.assertTrue(has_layer_col)
        self.assertEqual(["SAPATA S1"], data["ITEM"].tolist())
        self.assertNotIn("AltoQi_Eberick_Elemento / Elemento", data.columns)
        self.assertAlmostEqual(12.5, data.loc[0, "FORMA_m2"])
        self.assertEqual([], bitolas)

    def test_load_qto_detects_processed_aco_columns_case_insensitively(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "lista_processada.xlsx"
            pd.DataFrame(
                {
                    "ITEM": ["SAPATA S1"],
                    "NIVEL": ["BALDRAME"],
                    "CONCRETO_m3": [1.0],
                    "FORMA_m2": [2.0],
                    "aco_8.0mm_kg": [3.5],
                }
            ).to_excel(path, index=False)

            data, bitolas, has_layer_col = self.mod.load_qto(path, default_nivel="PAV")

        self.assertFalse(has_layer_col)
        self.assertEqual([8.0], bitolas)
        self.assertAlmostEqual(3.5, data.loc[0, "aco_8.0mm_kg"])


if __name__ == "__main__":
    unittest.main()
