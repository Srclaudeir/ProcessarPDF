import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path
import json
import os

from processar import (
    resource_path,
    carregar_schema_extracao,
    gerar_particoes_dinamicamente,
    extrair_texto_do_pdf,
    preencher_excel_novo_com_placeholders,
    salvar_texto_em_arquivo,
    salvar_json_em_arquivo,
    achatar_json,
)

class TestProcessar(unittest.TestCase):

    @patch("processar.Path")
    def test_resource_path_dev_mode(self, mock_path):
        mock_path.return_value = Path("/mocked/path")
        result = resource_path("test_file.txt")
        self.assertEqual(result, Path(__file__).resolve().parent / "test_file.txt")

    @patch("processar.Path")
    @patch("processar.sys")
    def test_resource_path_pyinstaller_mode(self, mock_sys, mock_path):
        mock_sys._MEIPASS = "/mocked/meipass"
        mock_path.return_value = Path("/mocked/path")
        result = resource_path("test_file.txt")
        self.assertEqual(result, Path("/mocked/meipass/test_file.txt"))

    @patch("processar.open", new_callable=MagicMock)
    @patch("processar.resource_path")
    def test_carregar_schema_extracao_valid(self, mock_resource_path, mock_open):
        mock_resource_path.return_value = Path("/mocked/schema.json")
        mock_open.return_value.__enter__.return_value.read.return_value = json.dumps({
            "bloco1": {"json_chave": "chave1", "particao": 1},
            "bloco2": {"json_chave": "chave2", "particao": 2}
        })
        result = carregar_schema_extracao()
        self.assertTrue(result)

    @patch("processar.open", new_callable=MagicMock)
    @patch("processar.resource_path")
    def test_carregar_schema_extracao_invalid(self, mock_resource_path, mock_open):
        mock_resource_path.return_value = Path("/mocked/schema.json")
        mock_open.return_value.__enter__.return_value.read.return_value = "invalid json"
        result = carregar_schema_extracao()
        self.assertFalse(result)

    @patch("processar.BLOCO_CONFIG", {"bloco1": {"particao": 1}, "bloco2": {"particao": 2}})
    def test_gerar_particoes_dinamicamente(self):
        result = gerar_particoes_dinamicamente()
        self.assertTrue(result)

    @patch("processar.pdfplumber.open")
    def test_extrair_texto_do_pdf(self, mock_pdfplumber_open):
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock(), MagicMock()]
        mock_pdf.pages[0].extract_text.return_value = "Texto da página 1"
        mock_pdf.pages[1].extract_text.return_value = "Texto da página 2"
        mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

        result = extrair_texto_do_pdf(Path("/mocked/file.pdf"))
        self.assertEqual(result, "Texto da página 1\nTexto da página 2")

    @patch("processar.openpyxl.load_workbook")
    @patch("processar.open")
    def test_preencher_excel_novo_com_placeholders(self, mock_open, mock_load_workbook):
        mock_open.return_value.__enter__.return_value.read.return_value = json.dumps({
            "{{PLACEHOLDER1}}": "Valor1",
            "{{PLACEHOLDER2}}": "Valor2"
        })
        mock_workbook = MagicMock()
        mock_sheet = MagicMock()
        mock_sheet.iter_rows.return_value = [[MagicMock(value="{{PLACEHOLDER1}}")], [MagicMock(value="{{PLACEHOLDER2}}")]]
        mock_workbook.active = mock_sheet
        mock_load_workbook.return_value = mock_workbook

        result = preencher_excel_novo_com_placeholders(
            Path("/mocked/json.json"),
            Path("/mocked/template.xlsx"),
            Path("/mocked/output.xlsx")
        )
        self.assertTrue(result)

    @patch("processar.open")
    def test_salvar_texto_em_arquivo(self, mock_open):
        salvar_texto_em_arquivo("Texto de teste", Path("/mocked/file.txt"))
        mock_open.assert_called_once_with(Path("/mocked/file.txt"), "w", encoding="utf-8")
        mock_open.return_value.__enter__.return_value.write.assert_called_once_with("Texto de teste")

    @patch("processar.open")
    def test_salvar_json_em_arquivo(self, mock_open):
        salvar_json_em_arquivo({"key": "value"}, Path("/mocked/file.json"))
        mock_open.assert_called_once_with(Path("/mocked/file.json"), "w", encoding="utf-8")
        mock_open.return_value.__enter__.return_value.write.assert_called_once_with(json.dumps({"key": "value"}, ensure_ascii=False, indent=4))

    def test_achatar_json(self):
        input_json = {
            "key1": {"subkey1": "value1", "subkey2": "value2"},
            "key2": ["item1", {"subitem1": "value3"}]
        }
        expected_output = {
            "key1_subkey1": "value1",
            "key1_subkey2": "value2",
            "key2_1": "item1",
            "key2_2_subitem1": "value3"
        }
        result = achatar_json(input_json)
        self.assertEqual(result, expected_output)

if __name__ == "__main__":
    unittest.main()