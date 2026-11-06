"""Leitura de CSV e exportação de planilhas, sem dependência da interface."""
from __future__ import annotations

import codecs
import csv
import io
import os
import tempfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from charset_normalizer import from_bytes


@dataclass(frozen=True)
class CsvInfo:
    encoding: str
    delimiter: str
    quotechar: str


def decode_csv(data: bytes, encoding: str | None = None) -> tuple[str, str]:
    if encoding:
        return data.decode(encoding), encoding
    # BOMs devem ser verificados antes de UTF-8; UTF-32 começa como UTF-16.
    for bom, name in ((codecs.BOM_UTF32_LE, 'utf-32'),
                      (codecs.BOM_UTF32_BE, 'utf-32'),
                      (codecs.BOM_UTF8, 'utf-8-sig'),
                      (codecs.BOM_UTF16_LE, 'utf-16'),
                      (codecs.BOM_UTF16_BE, 'utf-16')):
        if data.startswith(bom):
            return data.decode(name), name
    try:
        return data.decode('utf-8'), 'utf-8'
    except UnicodeDecodeError:
        pass
    # Windows-1252 é muito comum nos CSVs brasileiros. Use a análise estatística
    # para outros alfabetos, preferindo cp1252 quando a detecção for ambígua.
    guess = from_bytes(data).best()
    if guess is not None and guess.encoding not in {'cp1250', 'cp1252', 'cp1254', 'cp1257', 'iso8859_1', 'iso8859_15'}:
        return str(guess), guess.encoding
    try:
        return data.decode('cp1252'), 'cp1252'
    except UnicodeDecodeError:
        if guess is not None:
            return str(guess), guess.encoding
        return data.decode('latin-1'), 'latin-1'


def detect_format(text: str, encoding: str) -> CsvInfo:
    sample = text[:131072]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|:')
        return CsvInfo(encoding, dialect.delimiter, dialect.quotechar)
    except csv.Error:
        # Também permite CSV de uma coluna, inclusive com campos entre aspas.
        counts = Counter(char for char in sample.split('\n', 1)[0] if char in ',;\t|:')
        delimiter = counts.most_common(1)[0][0] if counts else ','
        quote = "'" if sample.lstrip().startswith("'") else '"'
        return CsvInfo(encoding, delimiter, quote)


def read_csv(path: str | Path, encoding: str | None = None,
             delimiter: str | None = None, quotechar: str | None = None
             ) -> tuple[list[list[str]], CsvInfo]:
    text, detected_encoding = decode_csv(Path(path).read_bytes(), encoding)
    info = detect_format(text, detected_encoding)
    info = CsvInfo(info.encoding, delimiter or info.delimiter, quotechar or info.quotechar)
    if not text.strip():
        raise ValueError('O arquivo CSV está vazio.')
    rows = list(csv.reader(io.StringIO(text, newline=''), delimiter=info.delimiter,
                           quotechar=info.quotechar, strict=True))
    return rows, info


def export_rows(rows: list[list[str]], destination: str | Path) -> None:
    """Grava atomicamente. Mantém texto, zeros iniciais e fórmulas como texto."""
    destination = Path(destination)
    suffix = destination.suffix.lower()
    if suffix not in {'.xlsx', '.ods'}:
        raise ValueError('Escolha um arquivo .xlsx ou .ods.')
    if suffix == '.xlsx' and (len(rows) > 1_048_576 or any(len(r) > 16_384 for r in rows)):
        raise ValueError('O CSV excede o limite de linhas ou colunas do XLSX.')
    if suffix == '.xlsx' and any(len(value) > 32767 for row in rows for value in row):
        raise ValueError('Uma célula excede o limite de 32.767 caracteres do XLSX.')
    handle, temp = tempfile.mkstemp(suffix=suffix, dir=destination.parent)
    os.close(handle)
    try:
        if suffix == '.xlsx':
            from openpyxl import Workbook
            from openpyxl.cell import WriteOnlyCell
            workbook = Workbook(write_only=True)
            sheet = workbook.create_sheet('Dados')
            for row in rows:
                cells = []
                for value in row:
                    cell = WriteOnlyCell(sheet, value=value)
                    cell.data_type = 's'
                    cells.append(cell)
                sheet.append(cells)
            workbook.save(temp)
            workbook.close()
        else:
            from odf.opendocument import OpenDocumentSpreadsheet
            from odf.table import Table, TableCell, TableRow
            from odf.text import P
            from odf import teletype
            workbook = OpenDocumentSpreadsheet()
            sheet = Table(name='Dados')
            for row in rows:
                output_row = TableRow()
                for value in row:
                    cell = TableCell(valuetype='string')
                    paragraph = P()
                    teletype.addTextToElement(paragraph, value)
                    cell.addElement(paragraph)
                    output_row.addElement(cell)
                sheet.addElement(output_row)
            workbook.spreadsheet.addElement(sheet)
            workbook.save(temp, addsuffix=False)
        os.replace(temp, destination)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)
