"""Генератор JSON и SQL. Python 3.10+, Windows / Linux."""
import csv
import io
import json
import math
import os
import re
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

GROUP_NAME = 'INSERT INTO material.material_group_tnved'
TNVED_NAME = 'INSERT INTO material.material_tnved'


def read_rows(path, sheet=0):
    path = Path(path)
    if path.suffix.lower() == '.csv':
        raw = path.read_bytes()
        try:
            content = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            content = raw.decode('cp1251')
        try:
            dialect = csv.Sniffer().sniff(content[:65536], delimiters=';,\t,')
        except csv.Error:
            raise ValueError('Не удалось определить разделитель CSV. Используйте ; , или табуляцию.')
        return list(csv.reader(io.StringIO(content), dialect))
    if path.suffix.lower() == '.xlsx':
        import openpyxl
        book = openpyxl.load_workbook(path, read_only=True, data_only=False)
        try:
            return list(book.worksheets[sheet].iter_rows(values_only=True))
        finally:
            book.close()
    if path.suffix.lower() == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path))
        try:
            tab = book.sheet_by_index(sheet)
            return [tab.row_values(i) for i in range(tab.nrows)]
        finally:
            book.release_resources()
    raise ValueError('Поддерживаются только XLS, XLSX, CSV.')


def normalize(value, tnved=False):
    if value is None or str(value).strip() == '':
        raise ValueError('пустой код')
    if isinstance(value, bool):
        raise ValueError('логическое значение вместо кода')
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            raise ValueError('код должен быть целым числом')
        value = str(int(value))
    else:
        value = str(value).strip()
    if not re.fullmatch(r'[0-9]+', value):
        raise ValueError('код должен содержать только цифры; формулы и научная запись не принимаются')
    if tnved:
        if len(value) > 10:
            raise ValueError('ТН ВЭД длиннее 10 цифр')
        value = value.zfill(10)
    return value


def collect(rows, header=True):
    pairs = []
    seen = set()
    errors = []
    duplicates = 0
    for number, row in enumerate(rows, 1):
        if header and number == 1:
            continue
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        try:
            group = normalize(row[0] if row else None)
            tnved = normalize(row[1] if len(row) > 1 else None, tnved=True)
        except ValueError as exc:
            errors.append(f'Строка {number}: {exc}')
            continue
        pair = (group, tnved)
        if pair in seen:
            duplicates += 1
        else:
            seen.add(pair)
            pairs.append(pair)
    if errors:
        raise ValueError('Файлы не созданы. Исправьте входные данные:\n' + '\n'.join(errors[:25]) +
                         (f'\nВсего ошибок: {len(errors)}' if len(errors) > 25 else ''))
    if not pairs:
        raise ValueError('Нет заполненных связок. Проверьте файл и настройку заголовка.')
    return pairs, duplicates


def queries(pairs):
    groups = [f"INSERT INTO material.material_group_tnved (group_code, tnved_code, creation_date, update_date) VALUES ('{g}', '{t}', NOW(), NOW()) ON CONFLICT DO NOTHING;" for g, t in pairs]
    codes = list(dict.fromkeys(t for _, t in pairs))
    tnved_queries = []
    for start in range(0, len(codes), 1000):
        values = ',\n'.join(f"('{t}')" for t in codes[start:start + 1000])
        tnved_queries.append("INSERT INTO material.material_tnved (tnved_code, tnved_name, creation_date, update_date)\n"
            "SELECT v.tnved_code, ' ', NOW(), NOW()\nFROM (VALUES\n" + values +
            ") AS v(tnved_code)\nWHERE NOT EXISTS (SELECT 1 FROM material.material_tnved t WHERE t.tnved_code = v.tnved_code);")
    return groups, tnved_queries, len(codes)


def filename(value, format="json"):
    value = value.strip()
    if not value or re.search(r'[<>:"/\\|?*\x00-\x1f]', value) or value.endswith(('.', ' ')):
        raise ValueError('Укажите корректные имена файлов без пути и символов <>:"/\\|?*.')
    if value.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1,10)), *(f'LPT{i}' for i in range(1,10))}:
        raise ValueError('Это имя зарезервировано Windows.')
    if value.lower().endswith(('.json', '.sql')):
        value = value.rsplit('.', 1)[0]
    return value + '.' + format


TNVED_HEADER = '-- 1) Справочник ТН ВЭД (добавляются только отсутствующие коды)'
GROUP_HEADER = '-- 2) Привязка групп к ТН ВЭД'


def render_outputs(pairs, format='json', single=False):
    if format not in ('json', 'sql'):
        raise ValueError('Неизвестный формат вывода.')
    groups, codes, count = queries(pairs)
    if format == 'json':
        payloads = [codes + groups] if single else [groups, codes]
        return [json.dumps(data, ensure_ascii=False, indent=2) + '\n' for data in payloads], count
    dictionary = TNVED_HEADER + '\n' + '\n'.join(codes)
    links = GROUP_HEADER + '\n' + '\n'.join(groups)
    return ([dictionary + '\n\n' + links + '\n'] if single else [links + '\n', dictionary + '\n']), count


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ТН ВЭД → JSON / SQL · 1.1.0')
        self.geometry('900x590')
        self.minsize(800, 570)
        self.source = tk.StringVar()
        self.output = tk.StringVar(value=str(Path(sys.executable if getattr(sys, 'frozen', False) else __file__).resolve().parent))
        self.group_name = tk.StringVar(value=GROUP_NAME)
        self.tnved_name = tk.StringVar(value=TNVED_NAME)
        self.combined_name = tk.StringVar(value='tnved_and_group_tnved')
        self.format = tk.StringVar(value='json')
        self.layout = tk.StringVar(value='Два файла')
        self.header = tk.BooleanVar(value=True)
        self.sheet = tk.StringVar(value='1')
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)
        self.name_entries = []
        for index, (label, variable, button) in enumerate([
            ('Входной файл:', self.source, self.choose_source),
            ('Папка результата:', self.output, self.choose_output),
            ('Файл со связками:', self.group_name, None),
            ('Файл с ТН ВЭД:', self.tnved_name, None),
            ('Объединённый файл:', self.combined_name, None),
        ]):
            ttk.Label(frame, text=label).grid(row=index, column=0, sticky='w', pady=7)
            entry = ttk.Entry(frame, textvariable=variable)
            entry.grid(row=index, column=1, sticky='ew', padx=8)
            if index >= 2:
                self.name_entries.append(entry)
            if button:
                ttk.Button(frame, text='Выбрать…', command=button).grid(row=index, column=2)
        ttk.Label(frame, text='Формат вывода:').grid(row=5, column=0, sticky='w', pady=8)
        ttk.Combobox(frame, textvariable=self.format, values=('json', 'sql'), state='readonly', width=12).grid(row=5, column=1, sticky='w', padx=8)
        ttk.Label(frame, text='Количество файлов:').grid(row=6, column=0, sticky='w', pady=8)
        choice = ttk.Combobox(frame, textvariable=self.layout, values=('Два файла', 'Один файл'), state='readonly', width=16)
        choice.grid(row=6, column=1, sticky='w', padx=8)
        choice.bind('<<ComboboxSelected>>', self.update_layout)
        ttk.Checkbutton(frame, text='Первая строка – заголовки (пропустить)', variable=self.header).grid(row=7, column=1, sticky='w', pady=8)
        ttk.Label(frame, text='Номер листа Excel:').grid(row=8, column=0, sticky='w')
        ttk.Spinbox(frame, from_=1, to=999, textvariable=self.sheet, width=6).grid(row=8, column=1, sticky='w', padx=8)
        ttk.Label(frame, text='1-я колонка: группа · 2-я: ТН ВЭД. Остальные колонки игнорируются.\nТН ВЭД короче 10 цифр дополняется нулями слева.\nJSON: массив SQL-строк. SQL: готовый текст запросов.\nВ одном файле справочник идёт перед связками. Расширение добавляется автоматически.', wraplength=800).grid(row=9, column=0, columnspan=3, sticky='w', pady=16)
        ttk.Button(frame, text='Создать файлы', command=self.run).grid(row=10, column=1, sticky='ew', padx=8)
        self.update_layout()

    def update_layout(self, event=None):
        single = self.layout.get() == 'Один файл'
        for entry in self.name_entries[:2]:
            entry.configure(state='disabled' if single else 'normal')
        self.name_entries[2].configure(state='normal' if single else 'disabled')

    def choose_source(self):
        path = filedialog.askopenfilename(initialdir=self.output.get(), filetypes=[('Таблицы', '*.xls *.xlsx *.csv')])
        if path:
            self.source.set(path)
            self.output.set(str(Path(path).parent))

    def choose_output(self):
        path = filedialog.askdirectory(initialdir=self.output.get())
        if path:
            self.output.set(path)

    def run(self):
        try:
            source = Path(self.source.get())
            if not source.is_file():
                raise ValueError('Выберите существующий входной файл.')
            out = Path(self.output.get())
            if not out.is_dir():
                raise ValueError('Выберите существующую папку результата.')
            single = self.layout.get() == 'Один файл'
            names = [self.combined_name.get()] if single else [self.group_name.get(), self.tnved_name.get()]
            targets = [out / filename(name, self.format.get()) for name in names]
            if len({str(path).casefold() for path in targets}) != len(targets):
                raise ValueError('Имена файлов должны различаться.')
            sheet = int(self.sheet.get()) - 1
            if sheet < 0:
                raise ValueError('Номер листа начинается с 1.')
            pairs, duplicates = collect(read_rows(source, sheet), self.header.get())
            payloads, count = render_outputs(pairs, self.format.get(), single)
            existing = [str(p) for p in targets if p.exists()]
            if existing and not messagebox.askyesno('Перезапись', 'Заменить существующие файлы?\n' + '\n'.join(existing)):
                return
            # Prepare both files fully before replacing either destination.
            import tempfile
            staged = []
            try:
                for data in payloads:
                    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=out, delete=False, suffix='.tmp') as f:
                        staged.append(Path(f.name))
                        f.write(data)
                for temporary, target in zip(staged, targets):
                    os.replace(temporary, target)
            finally:
                for temporary in staged:
                    temporary.unlink(missing_ok=True)
            messagebox.showinfo('Готово', f'Создано связок: {len(pairs)}\nУникальных ТН ВЭД: {count}\nПовторов удалено: {duplicates}\n\n' + '\n'.join(map(str, targets)))
        except Exception as exc:
            messagebox.showerror('Ошибка', str(exc))


if __name__ == '__main__':
    App().mainloop()
