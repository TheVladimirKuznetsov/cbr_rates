import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import pandas as pd
from datetime import datetime
from openpyxl.utils import get_column_letter

CODES = {
    "USD (Доллар США)": "R01235",
    "EUR (Евро)": "R01239",
    "CNY (Китайский юань)": "R01375",
    "INR (Индийская рупия)": "R01270",
    "AED (Дирхам ОАЭ)": "R01230",
    "AUD (Австралийский доллар)": "R01010",
    "ATS (Австрийский шиллинг)": "R01015",
    "AZN (Азербайджанский манат)": "R01020",
    "DZD (Алжирский динар)": "R01030",
    "GBP (Фунт стерлингов)": "R01035",
    "AON (Ангольская новая кванза)": "R01040",
    "AMD (Армянский драм)": "R01060",
    "BHD (Бахрейнский динар)": "R01080",
    "BYN (Белорусский рубль)": "R01090",
    "BEF (Бельгийский франк)": "R01095",
    "BGN (Болгарский лев)": "R01100",
    "BOB (Боливиано)": "R01105",
    "BRL (Бразильский реал)": "R01115",
    "HUF (Венгерский форинт)": "R01135",
    "VND (Вьетнамский донг)": "R01150",
    "HKD (Гонконгский доллар)": "R01200",
    "GRD (Греческая драхма)": "R01205",
    "GEL (Грузинский лари)": "R01210",
    "DKK (Датская крона)": "R01215",
    "EGP (Египетский фунт)": "R01240",
    "IDR (Индонезийская рупия)": "R01280",
    "IRR (Иранский риал)": "R01300",
    "IEP (Ирландский фунт)": "R01305",
    "ISK (Исландская крона)": "R01310",
    "ESP (Испанская песета)": "R01315",
    "ITL (Итальянская лира)": "R01325",
    "KZT (Казахстанский тенге)": "R01335",
    "CAD (Канадский доллар)": "R01350",
    "QAR (Катарский риал)": "R01355",
    "KGS (Киргизский сом)": "R01370",
    "KWD (Кувейтский динар)": "R01390",
    "CUP (Кубинское песо)": "R01395",
    "LVL (Латвийский лат)": "R01405",
    "LBP (Ливанский фунт)": "R01420",
    "LTL (Литовский лит)": "R01435",
    "MDL (Молдавский лей)": "R01500",
    "MNT (Монгольский тугрик)": "R01503",
    "DEM (Немецкая марка)": "R01510",
    "NGN (Нигерийская найра)": "R01520",
    "NLG (Нидерландский гульден)": "R01523",
    "NZD (Новозеландский доллар)": "R01530",
    "NOK (Норвежская крона)": "R01535",
    "OMR (Оманский риал)": "R01540",
    "PLN (Польский злотый)": "R01565",
    "PTE (Португальский эскудо)": "R01570",
    "SAR (Саудовский риял)": "R01580",
    "RON (Румынский лей)": "R01585",
    "SGD (Сингапурский доллар)": "R01625",
    "SRD (Суринамский доллар)": "R01665",
    "TJS (Таджикский сомони)": "R01670",
    "THB (Тайский бат)": "R01675",
    "BDT (Бангладешская така)": "R01685",
    "TRY (Турецкая лира)": "R01700",
    "TMT (Туркменский новый манат)": "R01710",
    "UZS (Узбекский сум)": "R01717",
    "UAH (Украинская гривна)": "R01720",
    "FIM (Финляндская марка)": "R01740",
    "FRF (Французский франк)": "R01750",
    "CZK (Чешская крона)": "R01760",
    "SEK (Шведская крона)": "R01770",
    "CHF (Швейцарский франк)": "R01775",
    "EEK (Эстонская крона)": "R01795",
    "ETB (Эфиопский быр)": "R01800",
    "RSD (Сербский динар)": "R01804",
    "ZAR (Южноафриканский рэнд)": "R01810",
    "KRW (Южнокорейская вона)": "R01815",
    "JPY (Японская иена)": "R01820",
    "MMK (Мьянмский кьят)": "R02005",
}

# Функция загрузки одной валюты
def load_currency(code_id: str, label_display: str, date_from: str, date_to: str) -> pd.DataFrame:
    """
    Скачивает ежедневный курс ЦБ для валюты code_id в периоде [date_from, date_to]
    и возвращает DataFrame(Date, <короткий код>), где короткий код — до пробела.
    """
    url = (
        "https://www.cbr.ru/scripts/XML_dynamic.asp"
        f"?VAL_NM_RQ={code_id}&date_req1={date_from}&date_req2={date_to}"
    )
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    df = pd.read_xml(r.content, xpath=".//Record")
    if df is None or df.empty:
        # Вернём пустой с нужными колонками
        col = label_display.split()[0]  # "USD (Доллар...)" -> "USD"
        return pd.DataFrame(columns=["Date", col])

    df["Date"] = pd.to_datetime(df["Date"], format="%d.%m.%Y")
    col = label_display.split()[0]  # колонка в итоговой таблице: USD/EUR/...
    # Значение приходит с запятой, делим на Nominal (на случай номинала != 1)
    df[col] = (
        df["Value"].astype(str).str.replace(",", ".", regex=False).astype(float)
        / df["Nominal"].astype(float)
    )
    return df[["Date", col]]

# Приложение
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Курсы ЦБ РФ — выбор валют и периода")
        self.geometry("960x560")
        self.minsize(820, 520)

        # Верхняя панель управления
        control = ttk.Frame(self, padding=12)
        control.pack(side=tk.TOP, fill=tk.X)

        # Левый блок — список валют
        left = ttk.Frame(control)
        left.pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(left, text="Выберите валюты:").pack(anchor="w")
        self.lst = tk.Listbox(left, selectmode=tk.EXTENDED, height=12, exportselection=False)
        for k in CODES.keys():
            self.lst.insert(tk.END, k)
        self.lst.pack(fill=tk.BOTH, expand=True)

        # Средний блок — даты
        mid = ttk.Frame(control)
        mid.pack(side=tk.LEFT, padx=(0, 16), fill=tk.Y)
        ttk.Label(mid, text="Дата от (дд/мм/гггг):").grid(row=0, column=0, sticky="w")
        self.ent_from = ttk.Entry(mid, width=14)
        self.ent_from.grid(row=0, column=1, padx=(6, 0), sticky="w")
        self.ent_from.insert(0, "01/01/2000")

        ttk.Label(mid, text="Дата до (дд/мм/гггг):").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.ent_to = ttk.Entry(mid, width=14)
        self.ent_to.grid(row=1, column=1, padx=(6, 0), sticky="w", pady=(6, 0))
        self.ent_to.insert(0, "31/12/2025")

        # Правый блок — кнопки
        right = ttk.Frame(control)
        right.pack(side=tk.LEFT)
        self.btn_load = ttk.Button(right, text="Загрузить", command=self.on_load_clicked)
        self.btn_load.grid(row=0, column=0, sticky="ew")
        self.btn_export = ttk.Button(right, text="Экспорт в EXCEL", command=self.export_excel, state=tk.DISABLED)
        self.btn_export.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # Статусная строка
        self.status = ttk.Label(self, text="Готово", anchor="w", padding=(12, 6))
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # Таблица результатов
        table_frame = ttk.Frame(self, padding=(12, 0, 12, 12))
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame, show="headings")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.df_data: pd.DataFrame | None = None
        self.setup_empty_table()

    # ---- Вспомогательные методы GUI ----
    def setup_empty_table(self):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
            self.tree.column(col, width=80)
        self.tree["columns"] = ("Date",)
        self.tree.heading("Date", text="Date")
        self.tree.column("Date", width=120, anchor="w")

    def set_status(self, txt: str):
        self.status.config(text=txt)
        self.status.update_idletasks()

    def validate_date(self, s: str) -> bool:
        try:
            datetime.strptime(s, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    # ---- Обработчики ----
    def on_load_clicked(self):
        sel_indices = self.lst.curselection()
        if not sel_indices:
            messagebox.showwarning("Выбор валют", "Выберите хотя бы одну валюту.")
            return
        chosen = [self.lst.get(i) for i in sel_indices]

        dfrom = self.ent_from.get().strip()
        dto = self.ent_to.get().strip()
        if not (self.validate_date(dfrom) and self.validate_date(dto)):
            messagebox.showerror("Неверная дата", "Введите даты в формате дд/мм/гггг.")
            return

        self.btn_load.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.set_status("Загрузка…")

        threading.Thread(target=self._load_data_thread, args=(chosen, dfrom, dto), daemon=True).start()

    def _load_data_thread(self, chosen: list[str], dfrom: str, dto: str):
        try:
            dfs = []
            for label_display in chosen:
                code = CODES[label_display]
                df = load_currency(code, label_display, dfrom, dto)  # колонка: "USD", "EUR", ...
                dfs.append(df)

            if not dfs:
                raise RuntimeError("Нет данных.")
            data = dfs[0]
            for df in dfs[1:]:
                data = pd.merge(data, df, on="Date", how="outer")

            data = data.sort_values("Date").reset_index(drop=True)
            self.df_data = data

            self.after(0, self._populate_table)
            self.after(0, lambda: self.set_status(f"Готово. Строк: {len(data)}"))
            self.after(0, lambda: self.btn_export.config(state=tk.NORMAL))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Ошибка загрузки", str(e)))
            self.after(0, lambda: self.set_status("Ошибка"))
        finally:
            self.after(0, lambda: self.btn_load.config(state=tk.NORMAL))

    def _populate_table(self):
        # очистка
        for row in self.tree.get_children():
            self.tree.delete(row)

        if self.df_data is None or self.df_data.empty:
            self.setup_empty_table()
            return

        cols = list(self.df_data.columns)
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c == "Date" else 110, anchor="e" if c != "Date" else "w")

        for _, row in self.df_data.iterrows():
            values = [row["Date"].strftime("%Y-%m-%d")] + [None]*(len(cols)-1)
            for i, c in enumerate(cols[1:], start=1):
                v = row[c]
                values[i] = "" if pd.isna(v) else f"{v:.6f}"
            self.tree.insert("", tk.END, values=values)

    # ---- Экспорт в Excel ----
    def export_excel(self):
        if self.df_data is None or self.df_data.empty:
            messagebox.showinfo("Экспорт", "Нет данных для экспорта.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel (*.xlsx)", "*.xlsx")],
            initialfile="cbr_rates.xlsx",
            title="Сохранить как…",
        )
        if not path:
            return

        df = self.df_data.copy()
        # Гарантируем корректный тип даты
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"]).dt.date

        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                # Округлим числовые колонки до 6 знаков
                num_cols = [c for c in df.columns if c != "Date"]
                df[num_cols] = df[num_cols].apply(pd.to_numeric, errors="coerce").round(6)
                df.to_excel(writer, index=False, sheet_name="rates")

                # Автоширина колонок
                ws = writer.sheets["rates"]
                for j, col in enumerate(df.columns, start=1):
                    max_len = max(
                        len(str(col)),
                        *(len(str(v)) for v in df[col].astype(str).tolist())
                    )
                    ws.column_dimensions[get_column_letter(j)].width = min(max_len + 2, 50)

            messagebox.showinfo("Экспорт", f"Сохранено: {path}")
        except Exception as e:
            messagebox.showerror("Экспорт в Excel", f"Ошибка: {e}")

if __name__ == "__main__":
    App().mainloop()
