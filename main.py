"""CSV → XLSX / ODS — interface Tkinter."""
import threading
import queue
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from converter import export_rows, read_csv


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CSV → XLSX / ODS')
        self.geometry('800x520')
        self.minsize(650, 450)
        self.path = tk.StringVar()
        self.output_format = tk.StringVar(value='XLSX')
        self.encoding = tk.StringVar(value='Automático')
        self.delimiter = tk.StringVar(value='Automático')
        self.quote = tk.StringVar(value='Automático')
        self.status = tk.StringVar(value='Selecione um CSV para começar.')
        self.events = queue.Queue()
        self.busy = False
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Conversor de CSV', font=('Segoe UI', 20, 'bold')).pack(anchor='w')
        ttk.Label(frame, text='Preserve acentos, campos e códigos ao gerar sua planilha.').pack(anchor='w', pady=(0, 16))
        file_row = ttk.Frame(frame)
        file_row.pack(fill='x')
        ttk.Entry(file_row, textvariable=self.path, state='readonly').pack(side='left', fill='x', expand=True)
        self.browse = ttk.Button(file_row, text='Selecionar CSV…', command=self.select_file)
        self.browse.pack(side='left', padx=(8, 0))
        options = ttk.Frame(frame)
        options.pack(fill='x', pady=16)
        self.combos = []
        for col, (label, variable, values) in enumerate([
            ('Saída', self.output_format, ['XLSX', 'ODS']),
            ('Codificação', self.encoding, ['Automático', 'utf-8-sig', 'cp1252', 'latin-1', 'utf-16', 'utf-32']),
            ('Delimitador', self.delimiter, ['Automático', ';', ',', 'Tabulação', '|', ':']),
            ('Aspas', self.quote, ['Automático', 'Duplas (")', "Simples (')"]),
        ]):
            ttk.Label(options, text=label).grid(row=0, column=col, sticky='w')
            combo = ttk.Combobox(options, textvariable=variable, values=values, state='readonly', width=17)
            combo.grid(row=1, column=col, padx=(0, 10), sticky='ew')
            options.columnconfigure(col, weight=1)
            self.combos.append(combo)
        actions = ttk.Frame(frame)
        actions.pack(fill='x', pady=(0, 12))
        self.preview_button = ttk.Button(actions, text='Analisar / atualizar prévia', command=self.preview)
        self.preview_button.pack(side='left')
        self.convert_button = ttk.Button(actions, text='Converter e salvar…', command=self.convert)
        self.convert_button.pack(side='right')
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill='both', expand=True)
        self.table = ttk.Treeview(table_frame, show='headings', height=8)
        vertical = ttk.Scrollbar(table_frame, orient='vertical', command=self.table.yview)
        horizontal = ttk.Scrollbar(table_frame, orient='horizontal', command=self.table.xview)
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.table.grid(row=0, column=0, sticky='nsew')
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal.grid(row=1, column=0, sticky='ew')
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        ttk.Label(frame, textvariable=self.status, wraplength=730).pack(anchor='w', pady=(12, 0))
        self.after(100, self.poll)

    def select_file(self):
        path = filedialog.askopenfilename(title='Selecionar CSV', filetypes=[('Arquivos CSV e texto', '*.csv *.txt *.tsv'), ('Todos os arquivos', '*.*')])
        if path:
            self.path.set(path)
            self.preview()

    def settings(self):
        return dict(encoding=None if self.encoding.get() == 'Automático' else self.encoding.get(),
                    delimiter=None if self.delimiter.get() == 'Automático' else ('\t' if self.delimiter.get() == 'Tabulação' else self.delimiter.get()),
                    quotechar=None if self.quote.get() == 'Automático' else ('"' if self.quote.get().startswith('Duplas') else "'"))

    def start(self, destination=None):
        if self.busy:
            return
        if not self.path.get():
            messagebox.showwarning('Selecione um arquivo', 'Selecione um CSV primeiro.')
            return
        source, settings = self.path.get(), self.settings()
        self.busy = True
        for widget in [self.browse, self.preview_button, self.convert_button, *self.combos]:
            widget.configure(state='disabled')
        self.status.set('Processando arquivo…')

        def work():
            try:
                rows, info = read_csv(source, **settings)
                if destination:
                    export_rows(rows, destination)
                self.events.put(('ok', (rows[:20], len(rows), info, destination)))
            except Exception as exc:
                self.events.put(('error', str(exc)))
        threading.Thread(target=work, daemon=True).start()

    def preview(self):
        self.start()

    def convert(self):
        if not self.path.get():
            messagebox.showwarning('Selecione um arquivo', 'Selecione um CSV primeiro.')
            return
        extension = '.' + self.output_format.get().lower()
        destination = filedialog.asksaveasfilename(title='Salvar planilha', defaultextension=extension,
            initialfile=Path(self.path.get()).stem + extension,
            filetypes=[(self.output_format.get(), '*' + extension)], confirmoverwrite=True)
        if destination:
            if Path(destination).suffix.lower() != extension:
                messagebox.showerror('Extensão inválida', f'Use a extensão {extension} para o formato selecionado.')
                return
            if Path(destination).resolve() == Path(self.path.get()).resolve():
                messagebox.showerror('Destino inválido', 'Escolha um destino diferente do arquivo de origem.')
                return
            self.start(destination)

    def poll(self):
        try:
            event, payload = self.events.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            for widget in [self.browse, self.preview_button, self.convert_button]:
                widget.configure(state='normal')
            for combo in self.combos:
                combo.configure(state='readonly')
            if event == 'error':
                self.status.set('Não foi possível processar o arquivo. Confira as opções de leitura.')
                messagebox.showerror('Erro na conversão', payload)
            else:
                rows, count, info, destination = payload
                self.table.delete(*self.table.get_children())
                columns = [str(i) for i in range(max((len(row) for row in rows), default=0))]
                self.table.configure(columns=columns)
                for i, column in enumerate(columns):
                    self.table.heading(column, text=f'Coluna {i + 1}')
                    self.table.column(column, width=140, stretch=False)
                for row in rows:
                    self.table.insert('', 'end', values=row)
                delimiter = 'Tabulação' if info.delimiter == '\t' else repr(info.delimiter)
                self.status.set(f'{count} registros • {info.encoding} • delimitador {delimiter} • aspas {info.quotechar} • prévia de até 20 registros')
                if destination:
                    messagebox.showinfo('Conversão concluída', f'Planilha salva em:\n{destination}')
        self.after(100, self.poll)


if __name__ == '__main__':
    Application().mainloop()
