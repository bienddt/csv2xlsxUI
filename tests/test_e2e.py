"""Fluxos completos com Tk real; somente diálogos do sistema são substituídos.

Executar em uma sessão gráfica: python -m unittest discover -s tests -p test_e2e.py -v
"""
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from odf import teletype
from odf.opendocument import load
from odf.table import TableCell, TableRow
from openpyxl import load_workbook

from main import Application


class GuiEndToEndTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.dialogs = {}
        for name in ('filedialog.askopenfilename', 'filedialog.asksaveasfilename',
                     'messagebox.showinfo', 'messagebox.showerror', 'messagebox.showwarning'):
            patcher = patch('main.' + name)
            self.dialogs[name] = patcher.start()
            self.addCleanup(patcher.stop)
        self.dialogs['filedialog.askopenfilename'].return_value = ''
        self.dialogs['filedialog.asksaveasfilename'].return_value = ''
        # A ausência de display deve falhar, em vez de mascarar E2E como aprovado.
        self.app = Application()
        self.addCleanup(self.close_app)
        self.callback_errors = []
        self.app.report_callback_exception = lambda *error: self.callback_errors.append(error)
        self.app.update()

    def close_app(self):
        self.wait_until(lambda: not self.app.busy)
        for timer in self.app.tk.splitlist(self.app.tk.call('after', 'info')):
            self.app.after_cancel(timer)
        self.app.destroy()

    def wait_until(self, condition, timeout=10):
        deadline = time.monotonic() + timeout
        while not condition():
            if time.monotonic() >= deadline:
                self.fail('Timeout aguardando a interface: ' + self.app.status.get())
            self.app.update()
            self.assertFalse(self.callback_errors, repr(self.callback_errors))
            time.sleep(0.01)
        self.app.update()
        self.assertFalse(self.callback_errors, repr(self.callback_errors))

    def select(self, content, encoding='utf-8'):
        source = self.root / 'entrada.csv'
        source.write_bytes(content.encode(encoding))
        self.dialogs['filedialog.askopenfilename'].return_value = str(source)
        self.app.browse.invoke()
        self.wait_until(lambda: not self.app.busy)
        self.assertEqual(self.app.path.get(), str(source))
        return source

    def preview_rows(self):
        return [list(self.app.table.item(item, 'values'))
                for item in self.app.table.get_children()]

    def choose_option(self, index, value):
        combo = self.app.combos[index]
        combo.set(value)
        combo.event_generate('<<ComboboxSelected>>')
        self.app.update()

    def save(self, extension):
        self.choose_option(0, extension.upper())
        destination = self.root / ('saida.' + extension)
        self.dialogs['filedialog.asksaveasfilename'].return_value = str(destination)
        self.app.convert_button.invoke()
        self.wait_until(lambda: not self.app.busy)
        self.dialogs['messagebox.showerror'].assert_not_called()
        self.dialogs['messagebox.showinfo'].assert_called_once()
        self.assertIn(str(destination), self.dialogs['messagebox.showinfo'].call_args.args[1])
        self.assertTrue(destination.is_file())
        self.assertEqual(self.dialogs['filedialog.asksaveasfilename'].call_args.kwargs['defaultextension'], '.' + extension)
        return destination

    def test_xlsx_utf8_preview_and_saved_values(self):
        source = self.select('código;nome;nota\n0012;João;"ação; coração\nsegunda linha"\n0013;Maria;=1+1\n', 'utf-8-sig')
        original = source.read_bytes()
        expected = [['código', 'nome', 'nota'], ['0012', 'João', 'ação; coração\nsegunda linha'], ['0013', 'Maria', '=1+1']]
        self.assertEqual(self.preview_rows(), expected)
        self.assertIn('utf-8-sig', self.app.status.get())
        destination = self.save('xlsx')
        book = load_workbook(destination)
        try:
            self.assertEqual(list(book.active.values), [tuple(row) for row in expected])
            self.assertEqual(book.active['C3'].data_type, 's')
        finally:
            book.close()
        self.assertEqual(source.read_bytes(), original)

    def test_ods_cp1252_single_quotes(self):
        self.select("código,nome,nota\n'0007','João','d''água, café'\n'0008','Ação','São Paulo'\n", 'cp1252')
        expected = [['código', 'nome', 'nota'], ['0007', 'João', "d'água, café"], ['0008', 'Ação', 'São Paulo']]
        self.assertEqual(self.preview_rows(), expected)
        self.assertIn('cp1252', self.app.status.get())
        document = load(self.save('ods'))
        actual = [[teletype.extractText(cell) for cell in row.getElementsByType(TableCell)]
                  for row in document.getElementsByType(TableRow)]
        self.assertEqual(actual, expected)

    def test_manual_options_and_reanalysis(self):
        self.select("'nome'|'nota'\n'João'|'ação|café'\n", 'cp1252')
        for index, value in [(1, 'cp1252'), (2, '|'), (3, "Simples (')")]:
            self.choose_option(index, value)
        self.app.preview_button.invoke()
        self.wait_until(lambda: not self.app.busy)
        self.assertEqual(self.preview_rows(), [['nome', 'nota'], ['João', 'ação|café']])
        self.save('xlsx')

    def test_preview_limit_does_not_truncate_export(self):
        self.select('nome;valor\n' + ''.join(f'Pessoa {i};00{i}\n' for i in range(35)))
        self.assertEqual(len(self.preview_rows()), 20)
        self.assertIn('36 registros', self.app.status.get())
        book = load_workbook(self.save('xlsx'))
        try:
            self.assertEqual(book.active.max_row, 36)
            self.assertEqual(book.active['B36'].value, '0034')
        finally:
            book.close()

    def test_cancel_selection_and_save(self):
        self.app.browse.invoke()
        self.assertEqual(self.app.path.get(), '')
        self.assertFalse(self.app.busy)
        self.select('a;b\n1;2\n')
        self.app.convert_button.invoke()
        self.assertFalse(self.app.busy)
        self.assertEqual(list(self.root.iterdir()), [self.root / 'entrada.csv'])
        self.dialogs['messagebox.showinfo'].assert_not_called()

    def test_conversion_without_source_shows_warning(self):
        self.app.convert_button.invoke()
        self.dialogs['messagebox.showwarning'].assert_called_once()
        self.dialogs['filedialog.asksaveasfilename'].assert_not_called()

    def test_wrong_extension_does_not_overwrite(self):
        self.select('a;b\n1;2\n')
        destination = self.root / 'existente.ods'
        destination.write_bytes(b'preservar')
        self.dialogs['filedialog.asksaveasfilename'].return_value = str(destination)
        self.app.convert_button.invoke()
        self.dialogs['messagebox.showerror'].assert_called_once()
        self.assertEqual(destination.read_bytes(), b'preservar')
        self.assertFalse(self.app.busy)

    def test_worker_error_recovers_and_allows_next_conversion(self):
        self.select('')
        self.dialogs['messagebox.showerror'].assert_called_once()
        for button in [self.app.browse, self.app.preview_button, self.app.convert_button]:
            self.assertNotIn('disabled', button.state())
        for combo in self.app.combos:
            self.assertIn('readonly', combo.state())
            self.assertNotIn('disabled', combo.state())
        self.dialogs['messagebox.showerror'].reset_mock()
        self.select('nome;cidade\nJoão;São Paulo\n')
        self.save('xlsx')

    def test_ui_remains_responsive_and_controls_disabled_during_work(self):
        import threading
        from converter import read_csv

        release = threading.Event()
        self.addCleanup(release.set)

        def slow_read(*args, **kwargs):
            if not release.wait(5):
                raise TimeoutError('O teste não liberou a leitura.')
            return read_csv(*args, **kwargs)

        source = self.root / 'entrada.csv'
        source.write_text('a;b\n1;2\n', encoding='utf-8')
        self.dialogs['filedialog.askopenfilename'].return_value = str(source)
        with patch('main.read_csv', side_effect=slow_read):
            self.app.browse.invoke()
            self.assertTrue(self.app.busy)
            for widget in [self.app.browse, self.app.preview_button, self.app.convert_button, *self.app.combos]:
                self.assertIn('disabled', widget.state())
            heartbeat = []
            self.app.after(0, lambda: heartbeat.append(True))
            self.wait_until(lambda: bool(heartbeat))
            self.assertTrue(self.app.busy)
            release.set()
            self.wait_until(lambda: not self.app.busy)
        self.assertEqual(self.preview_rows(), [['a', 'b'], ['1', '2']])


if __name__ == '__main__':
    unittest.main()
