# EZCAD Otomasyon Uygulaması (Flowchart Temelli Yeniden Yapılandırma)

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from pywinauto.application import Application
import configparser
import time
import psutil

class EZCADApp:
    def __init__(self, root):
        self.root = root
        self.config_file = "ezcad_config.ini"
        self.load_config()

        root.title("EZCAD - Excel & EZD Seçici")

        # EZCAD2.exe seçimi
        tk.Label(root, text="EZCAD2.exe Yolu:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.ezcad_exe_var = tk.StringVar(value=self.ezcad_exe_path or "")
        tk.Entry(root, textvariable=self.ezcad_exe_var, width=60).grid(row=0, column=1, padx=5)
        tk.Button(root, text="exe Seç", command=self.select_ezcad_exe).grid(row=0, column=2, padx=5)

        # Excel dosyası seçimi
        tk.Label(root, text="Excel (.xls) Dosyası:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.excel_path_var = tk.StringVar()
        tk.Entry(root, textvariable=self.excel_path_var, width=60).grid(row=1, column=1, padx=5)
        tk.Button(root, text="xls Seç", command=self.select_excel).grid(row=1, column=2, padx=5)

        # EZD dosyası seçimi
        tk.Label(root, text="EZCAD (.ezd) Dosyası:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.ezd_path_var = tk.StringVar()
        tk.Entry(root, textvariable=self.ezd_path_var, width=60).grid(row=2, column=1, padx=5)
        tk.Button(root, text="ezd Seç", command=self.select_ezd).grid(row=2, column=2, padx=5)

        # Excel önizleme
        tk.Label(root, text="Excel Önizleme (ilk 5 satır, 10 sütun):").grid(row=3, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 0))
        self.preview_box = tk.Text(root, height=8, width=90)
        self.preview_box.grid(row=4, column=0, columnspan=3, padx=10, pady=5)

        # Butonlar
        button_frame = tk.Frame(root)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)

        self.red_button = tk.Button(button_frame, text="Red", state=tk.DISABLED, command=self.send_red)
        self.red_button.pack(side=tk.LEFT, padx=10)

        self.mark_button = tk.Button(button_frame, text="Mark", state=tk.DISABLED, command=self.send_mark)
        self.mark_button.pack(side=tk.LEFT, padx=10)

        self.select_window_button = tk.Button(button_frame, text="EZCAD Penceresini Seç", command=self.select_ezcad_window)
        self.select_window_button.pack(side=tk.LEFT, padx=10)

        self.run_button = tk.Button(button_frame, text="Aktar (EZCAD Başlat)", command=self.run_ezcad)
        self.run_button.pack(side=tk.LEFT, padx=10)

    def select_ezcad_exe(self):
        file_path = filedialog.askopenfilename(title="EZCAD2.exe yolunu seçin", filetypes=[("EZCAD2", "EZCAD2.exe")])
        if file_path:
            self.ezcad_exe_var.set(file_path)
            self.save_config(file_path)

    def select_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            self.excel_path_var.set(file_path)
            try:
                df = pd.read_excel(file_path, engine='xlrd' if file_path.endswith('.xls') else None)
                preview_rows = []
                for row_idx in range(1, 6):
                    row_cells = []
                    for col_idx in range(1, 11):
                        try:
                            val = str(df.iat[row_idx, col_idx])
                        except:
                            val = ""
                        row_cells.append(val.ljust(8)[:8])
                    preview_rows.append(" ".join(row_cells))
                self.preview_box.delete("1.0", tk.END)
                self.preview_box.insert(tk.END, "\n".join(preview_rows))
            except Exception as e:
                self.preview_box.delete("1.0", tk.END)
                self.preview_box.insert(tk.END, f"Excel okunamadı: {e}")

    def select_ezd(self):
        file_path = filedialog.askopenfilename(filetypes=[("EZD files", "*.ezd")])
        if not file_path:
            return
        self.ezd_path_var.set(file_path)

    def run_ezcad(self):
        # self.root.withdraw()  # Arayüz gizleme kaldırıldı
        for proc in psutil.process_iter(['pid', 'name']):
            if 'EZCAD2.exe' in proc.info['name']:
                messagebox.showerror("Uyarı", "Lütfen önce açık olan EZCAD2 uygulamasını kapatın.")
                return

        exe_path = self.ezcad_exe_var.get()
        file_path = self.ezd_path_var.get()
        if not exe_path or not os.path.isfile(exe_path):
            messagebox.showerror("Hata", "Lütfen önce EZCAD2.exe yolunu girin.")
            return

        if not file_path:
            messagebox.showerror("Hata", ".ezd dosyası seçilmedi.")
            return

        try:
            os.spawnv(os.P_NOWAIT, exe_path, [exe_path, file_path])
            for _ in range(15):
                try:
                    app = Application(backend="win32").connect(title_re=".*License.*|.*Agreement.*|.*Terms.*")
                    license_win = app.top_window()
                    agree_button = license_win.child_window(title_re=".*agree.*|.*Kabul.*|.*Annehmen.*|.*同意.*", control_type="Button")
                    if agree_button.exists():
                        license_win.set_focus()
                        agree_button.click()
                        time.sleep(0.5)
                        license_win.minimize()
                        time.sleep(0.2)
                        license_win.type_keys('{ENTER}', set_foreground=False)
                        break
                except:
                    time.sleep(1)
        except Exception as e:
            messagebox.showerror("Başlatma Hatası", f"EZD dosyası başlatılamadı: {e}")
            return

        for _ in range(10):
            try:
                ezd_name = os.path.basename(file_path).replace(".", "\.")
                app = Application().connect(title_re=f".*{ezd_name}.*")
                self.ezcad_window = app.top_window()
                self.ezcad_window.minimize()
                break
            except:
                time.sleep(1)
        else:
            messagebox.showerror("Bağlantı Hatası", "EZCAD penceresi otomatik olarak algılanamadı.")
            return

        self.mark_button.config(state=tk.NORMAL)
        # self.root.deiconify()  # Arayüz geri getirme kaldırıldı
        self.red_button.config(state=tk.NORMAL)
        self.red_button.config(state=tk.NORMAL)

    def select_ezcad_window(self):
        try:
            file_path = self.ezd_path_var.get()
            if not file_path:
                raise Exception(".ezd dosyası seçilmedi.")
            ezd_name = os.path.basename(file_path).replace(".", "\.")
            app = Application().connect(title_re=f".*{ezd_name}.*")
            self.ezcad_window = app.top_window()
            messagebox.showinfo("Başarılı", ".ezd penceresi seçildi.")
        except Exception as e:
            messagebox.showerror("Hata", f"EZCAD penceresi bulunamadı: {e}")

    def send_red(self):
        try:
            if hasattr(self, 'ezcad_window'):
                self.ezcad_window.type_keys("{F1}")
            else:
                raise Exception("EZCAD penceresi tanımlı değil.")
        except Exception as e:
            messagebox.showerror("Hata", f"RED komutu gönderilemedi: {e}")

    def send_mark(self):
        try:
            if hasattr(self, 'ezcad_window'):
                self.ezcad_window.type_keys("{F2}")
            else:
                raise Exception("EZCAD penceresi tanımlı değil.")
        except Exception as e:
            messagebox.showerror("Hata", f"MARK komutu gönderilemedi: {e}")

    def load_config(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            config.read(self.config_file)
            self.ezcad_exe_path = config.get("Paths", "ezcad_exe", fallback="")
        else:
            self.ezcad_exe_path = ""

    def save_config(self, path):
        config = configparser.ConfigParser()
        config["Paths"] = {"ezcad_exe": path}
        with open(self.config_file, "w") as f:
            config.write(f)

if __name__ == "__main__":
    root = tk.Tk()
    app = EZCADApp(root)
    root.mainloop()