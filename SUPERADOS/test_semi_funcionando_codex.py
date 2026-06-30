import importlib.machinery
import importlib.util
import tempfile
import unittest
from pathlib import Path

import pandas as pd


def _load_module():
    path = Path(__file__).with_name("gerar_lista_materiais_SEMI FUNCIONANDO - CODEX.py")
    loader = importlib.machinery.SourceFileLoader("semi_codex", str(path))
    spec = importlib.util.spec_from_loader("semi_codex", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class SemiFuncionandoCodexTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.mod = _load_module()

    def test_load_qto_reads_description_columns_case_and_accent_insensitively(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "qto.xlsx"
            pd.DataFrame(
                {
                    "Ememento": ["S1"],
                    "Layer": ["BALDRAME"],
                    "descri\u00e7\u00e3o 1": ["Forma - Estrutura"],
                    "quantidade 1": ["12,5"],
                }
            ).to_excel(path, index=False)

            data, bitolas, has_layer_col = self.mod.load_qto(path, default_nivel="PAV")

        self.assertTrue(has_layer_col)
        self.assertEqual(["S1"], data["ELEM"].tolist())
        self.assertAlmostEqual(12.5, data.loc[0, "FORMA_m2"])
        self.assertEqual([], bitolas)

    def test_load_qto_detects_processed_aco_columns_case_insensitively(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "lista_processada.xlsx"
            pd.DataFrame(
                {
                    "ELEM": ["S1"],
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
