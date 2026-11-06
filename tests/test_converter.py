import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from odf.opendocument import load
from odf.table import TableCell, TableRow
from odf import teletype

from converter import export_rows, read_csv


class ConverterTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def read(self, text, source_encoding='utf-8', **options):
        source = self.root / 'input.csv'
        source.write_bytes(text.encode(source_encoding))
        return read_csv(source, **options)

    def test_utf8_double_quotes_multiline(self):
        rows, info = self.read('nome;observação\n"João";"ação; coração\nsegunda linha"\n')
        self.assertEqual(info.delimiter, ';')
        self.assertEqual(info.quotechar, '"')
        self.assertEqual(rows[1], ['João', 'ação; coração\nsegunda linha'])

    def test_single_quotes_and_escape(self):
        rows, info = self.read("nome,nota\n'Ana','d''água, café'\n")
        self.assertEqual(info.quotechar, "'")
        self.assertEqual(rows[1], ['Ana', "d'água, café"])

    def test_encodings(self):
        for encoding in ['utf-8-sig', 'utf-16', 'utf-32', 'cp1252']:
            with self.subTest(encoding=encoding):
                rows, _ = self.read('nome;cidade\nJoão;São Paulo\nAção;Brasília\n', encoding)
                self.assertEqual(rows[1], ['João', 'São Paulo'])

    def test_delimiters(self):
        for delimiter in [',', ';', '\t', '|', ':']:
            with self.subTest(delimiter=delimiter):
                rows, info = self.read(f'a{delimiter}b\n1{delimiter}2\n')
                self.assertEqual(info.delimiter, delimiter)
                self.assertEqual(rows[1], ['1', '2'])

    def test_manual_override(self):
        rows, info = self.read('ação|coração\n', 'cp1252', encoding='cp1252', delimiter='|', quotechar="'")
        self.assertEqual(rows, [['ação', 'coração']])
        self.assertEqual(info.encoding, 'cp1252')

    def test_empty_rejected(self):
        with self.assertRaises(ValueError):
            self.read('')

    def test_single_column(self):
        rows, _ = self.read('nome\nJoão\nMaria\n')
        self.assertEqual(rows, [['nome'], ['João'], ['Maria']])

    def test_exports_preserve_text(self):
        rows = [['código', 'texto', 'fórmula'], ['0012', 'ação  com\tespaços\ne linha', '=1+1']]
        xlsx = self.root / 'output.xlsx'
        export_rows(rows, xlsx)
        book = load_workbook(xlsx)
        self.assertEqual(list(book.active.values), [tuple(r) for r in rows])
        self.assertEqual(book.active['C2'].data_type, 's')
        book.close()
        ods = self.root / 'output.ods'
        export_rows(rows, ods)
        document = load(ods)
        actual = [[teletype.extractText(c) for c in row.getElementsByType(TableCell)]
                  for row in document.getElementsByType(TableRow)]
        self.assertEqual(actual, rows)

    def test_failed_export_preserves_destination(self):
        destination = self.root / 'output.xlsx'
        destination.write_bytes(b'original')
        with self.assertRaises(Exception):
            export_rows([['illegal\x01character']], destination)
        self.assertEqual(destination.read_bytes(), b'original')
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), ['output.xlsx'])

    def test_long_cell_rejected(self):
        with self.assertRaises(ValueError):
            export_rows([['a' * 32768]], self.root / 'output.xlsx')


if __name__ == '__main__':
    unittest.main()
