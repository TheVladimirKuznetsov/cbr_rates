import threading
from datetime import datetime
from typing import Dict, List, Optional

import requests
import pandas as pd


from openpyxl.utils import get_column_letter


import tkinter as tk
from tkinter import ttk, messagebox, filedialog


SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "moex-cets-gui/1.0"})

BASE = "https://iss.moex.com/iss/engines/currency/markets/selt/boards/CETS/securities/{secid}/candles.json"
CANDLE_COLS = "begin,open,high,low,close,value,volume"

# Поддерживаемые таймфреймы
TF_SPEC = {
    "1m":  {"fetch_interval": 1,  "resample": None},
    "5m":  {"fetch_interval": 1,  "resample": "5T"},
    "15m": {"fetch_interval": 1,  "resample": "15T"},
    "30m": {"fetch_interval": 1,  "resample": "30T"},  # fallback 10m→30m
    "1h":  {"fetch_interval": 60, "resample": None},
    "1d":  {"fetch_interval": 24, "resample": None},
}

# Доступные инструменты
SECIDS: Dict[str, str] = {
    "CNY/RUB TOM": "CNYRUB_TOM",
    "BYN/RUB TOM": "BYNRUB_TOM",
    "KZT/RUB TOM": "KZTRUB_TOM",
    "AMD/RUB TOM": "AMDRUB_TOM",
}

def _fetch_candles_raw(secid: str, date_from: str, date_till: str, interval: int) -> pd.DataFrame:
    url = BASE.format(secid=secid)
    params = {
        "from": date_from,
        "till": date_till,
        "interval": interval,
        "iss.only": "candles",
        "candles.columns": CANDLE_COLS,
    }
    all_rows, start = [], 0
    while True:
        p = params.copy(); p["start"] = start
        r = SESSION.get(url, params=p, timeout=30)
        r.raise_for_status()
        j = r.json()
        rows = j.get("candles", {}).get("data", [])
        if not rows:
            break
        df = pd.DataFrame(rows, columns=j["candles"]["columns"])
        all_rows.append(df)
        start += len(rows)

    if not all_rows:
        return pd.DataFrame(columns=CANDLE_COLS.split(","))

    out = pd.concat(all_rows, ignore_index=True)
    out["begin"] = pd.to_datetime(out["begin"])
    for c in ["open","high","low","close","value","volume"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    return out.sort_values("begin").reset_index(drop=True)

def _resample_ohlcv(df: pd.DataFrame, freq: str) -> pd.DataFrame:
    if df.empty: return df
    dfi = df.set_index("begin").sort_index()
    agg = dfi.resample(freq).agg({
        "open":"first","high":"max","low":"min","close":"last",
        "value":"sum","volume":"sum"
    })
    agg = agg.dropna(subset=["open","high","low","close"])
    return agg.reset_index().rename(columns={"begin":"datetime"})

def fetch_candles(secid: str, tf: str, date_from: str, date_till: str) -> pd.DataFrame:
    if tf not in TF_SPEC:
        raise ValueError(f"Неподдерживаемый таймфрейм: {tf}")
    spec = TF_SPEC[tf]
    df = _fetch_candles_raw(secid, date_from, date_till, spec["fetch_interval"])
    # Fallback для 30m: агрегируем 10m → 30m
    if df.empty and tf == "30m":
        df10 = _fetch_candles_raw(secid, date_from, date_till, 10)
        if not df10.empty:
            return _resample_ohlcv(df10, "30T")
    if spec["resample"]:
        return _resample_ohlcv(df, spec["resample"])
    else:
        return df.rename(columns={"begin":"datetime"})[
            ["datetime","open","high","low","close","value","volume"]
        ]

def parse_ru_date(s: str) -> str:
    """ 'дд/мм/гггг' -> 'гггг-мм-дд' """
    dt = datetime.strptime(s, "%d/%m/%Y")
    return dt.strftime("%Y-%m-%d")

def autosize_openpyxl(ws, df: pd.DataFrame, max_width: int = 50):
    for j, col in enumerate(df.columns, start=1):
        texts = [str(col)] + [str(v) for v in df[col].tolist()]
        width = min(max(len(t) for t in texts) + 2, max_width)
        ws.column_dimensions[get_column_letter(j)].width = width

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MOEX CETS — свечи по валютным парам")
        self.geometry("1060x620")
        self.minsize(900, 560)

        # Верхняя панель
        control = ttk.Frame(self, padding=12)
        control.pack(side=tk.TOP, fill=tk.X)

        # Левый блок — список инструментов
        left = ttk.Frame(control)
        left.pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(left, text="Инструменты (CETS):").pack(anchor="w")
        self.lst = tk.Listbox(left, selectmode=tk.EXTENDED, height=10, exportselection=False)
        for k in SECIDS.keys():
            self.lst.insert(tk.END, k)
        self.lst.pack(fill=tk.BOTH, expand=True)

        # Средний блок — даты и ТФ
        mid = ttk.Frame(control)
        mid.pack(side=tk.LEFT, padx=(0, 16), fill=tk.Y)

        ttk.Label(mid, text="Дата от (дд/мм/гггг):").grid(row=0, column=0, sticky="w")
        self.ent_from = ttk.Entry(mid, width=14)
        self.ent_from.grid(row=0, column=1, padx=(6, 0), sticky="w")
        self.ent_from.insert(0, "01/01/2019")

        ttk.Label(mid, text="Дата до (дд/мм/гггг):").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.ent_to = ttk.Entry(mid, width=14)
        self.ent_to.grid(row=1, column=1, padx=(6, 0), sticky="w", pady=(6, 0))
        self.ent_to.insert(0, "01/01/2030")

        ttk.Label(mid, text="Таймфрейм:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.cmb_tf = ttk.Combobox(mid, width=12, state="readonly",
                                   values=["1m","5m","15m","30m","1h","1d"])
        self.cmb_tf.grid(row=2, column=1, padx=(6, 0), sticky="w", pady=(10, 0))
        self.cmb_tf.set("15m")

        # Правый блок — кнопки
        right = ttk.Frame(control)
        right.pack(side=tk.LEFT)
        self.btn_load = ttk.Button(right, text="Загрузить", command=self.on_load_clicked)
        self.btn_load.grid(row=0, column=0, sticky="ew")
        self.btn_export = ttk.Button(right, text="Экспорт в Excel", command=self.export_excel, state=tk.DISABLED)
        self.btn_export.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # Статус
        self.status = ttk.Label(self, text="Готово", anchor="w", padding=(12, 6))
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # Таблица
        table_frame = ttk.Frame(self, padding=(12, 0, 12, 12))
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(table_frame, show="headings")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)

        
        self.df_all: Optional[pd.DataFrame] = None
        self.per_secid: Dict[str, pd.DataFrame] = {}

        self.setup_empty_table()

    def setup_empty_table(self):
        self.tree["columns"] = ("datetime",)
        self.tree.heading("datetime", text="datetime")
        self.tree.column("datetime", width=160, anchor="w")

    def set_status(self, txt: str):
        self.status.config(text=txt)
        self.status.update_idletasks()

    def validate_date(self, s: str) -> bool:
        try:
            datetime.strptime(s, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    def on_load_clicked(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showwarning("Выбор инструментов", "Выберите хотя бы один инструмент.")
            return
        chosen_labels = [self.lst.get(i) for i in sel]

        dfrom = self.ent_from.get().strip()
        dto = self.ent_to.get().strip()
        if not (self.validate_date(dfrom) and self.validate_date(dto)):
            messagebox.showerror("Неверная дата", "Введите даты в формате дд/мм/гггг.")
            return

        tf = self.cmb_tf.get()
        if tf not in TF_SPEC:
            messagebox.showerror("Таймфрейм", "Выберите корректный таймфрейм.")
            return

        self.btn_load.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.set_status("Загрузка…")

        threading.Thread(
            target=self._load_data_thread,
            args=(chosen_labels, dfrom, dto, tf),
            daemon=True
        ).start()

    def _load_data_thread(self, chosen_labels: List[str], dfrom: str, dto: str, tf: str):
        try:
            date_from = parse_ru_date(dfrom)
            date_till = parse_ru_date(dto)

            collected = []
            per_secid_local: Dict[str, pd.DataFrame] = {}

            for label in chosen_labels:
                secid = SECIDS[label]
                df = fetch_candles(secid, tf, date_from, date_till)
                if df.empty:
                    continue
                df = df.copy()
                df.insert(0, "SECID", secid)
                per_secid_local[secid] = df[["datetime","open","high","low","close","value","volume"]].copy()
                collected.append(df)

            if not collected:
                raise RuntimeError("Нет данных по выбранным инструментам за указанный период.")

            df_all = pd.concat(collected, ignore_index=True)
            df_all.sort_values(["SECID", "datetime"], inplace=True)
            self.df_all = df_all.reset_index(drop=True)
            self.per_secid = per_secid_local

            self.after(0, self._populate_table)
            self.after(0, lambda: self.set_status(f"Готово. Строк: {len(self.df_all)}"))
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

        if self.df_all is None or self.df_all.empty:
            self.setup_empty_table()
            return

        cols = ["datetime", "SECID", "open", "high", "low", "close", "value", "volume"]
        self.tree["columns"] = cols
        for c in cols:
            w = 170 if c == "datetime" else (90 if c == "SECID" else 110)
            anchor = "w" if c in ("datetime","SECID") else "e"
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor=anchor)

        for _, row in self.df_all.iterrows():
            vals = [
                pd.to_datetime(row["datetime"]).strftime("%Y-%m-%d %H:%M:%S"),
                row["SECID"],
                *(("" if pd.isna(row[c]) else f"{row[c]:.6f}") for c in ["open","high","low","close","value","volume"])
            ]
            self.tree.insert("", tk.END, values=vals)

    def export_excel(self):
        if self.df_all is None or self.df_all.empty:
            messagebox.showinfo("Экспорт", "Нет данных для экспорта.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel (*.xlsx)", "*.xlsx")],
            initialfile="moex_cets_candles.xlsx",
            title="Сохранить как…",
        )
        if not path:
            return

        try:
            with pd.ExcelWriter(path, engine="openpyxl",
                                datetime_format="yyyy-mm-dd hh:mm:ss",
                                date_format="yyyy-mm-dd") as writer:
                # ALL
                df_all = self.df_all.copy()
                df_all.to_excel(writer, index=False, sheet_name="ALL")
                autosize_openpyxl(writer.sheets["ALL"], df_all)

                # Листы по каждому SECID
                for secid, df in self.per_secid.items():
                    sheet = secid[:31]
                    out = df.copy()
                    out.to_excel(writer, index=False, sheet_name=sheet)
                    autosize_openpyxl(writer.sheets[sheet], out)

            messagebox.showinfo("Экспорт", f"Сохранено: {path}")
        except Exception as e:
            messagebox.showerror("Экспорт в Excel", f"Ошибка: {e}")

if __name__ == "__main__":
    App().mainloop()