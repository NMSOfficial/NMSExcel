import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk

class GradyanButon(tk.Button):
    def __init__(self, parent, renk1, renk2, *args, **kwargs):
        tk.Button.__init__(self, parent, *args, **kwargs)
        self.renk1 = renk1
        self.renk2 = renk2
        self.bind("<Enter>", self.gradyan_efekti)
        self.bind("<Leave>", self.normal_renk)

    def gradyan_efekti(self, event):
        self.config(bg=self.renk2)

    def normal_renk(self, event):
        self.config(bg=self.renk1)

def excel_sec():
    global df, dosya_yolu
    dosya_yolu = filedialog.askopenfilename(title="Excel dosyasını seçin", filetypes=[("Excel dosyaları", "*.xlsx")])
    if dosya_yolu:
        df = pd.read_excel(dosya_yolu)
        barkod_no_sor()

def barkod_no_sor():
    barkod_no = simpledialog.askinteger("Barkod No", "Lütfen Barkod Numarasını Girin:")
    if barkod_no:
        düzenle_butonu.config(state="normal")

def satir_guncelle():
    global df
    barkod = simpledialog.askinteger("Barkod No", "Lütfen Barkod Numarasını Girin:")
    if barkod:
        satir_index = df[df['Barkod'] == barkod].index
        if not satir_index.empty:
            satir_index = satir_index[0]
            satir_verileri = df.loc[satir_index].to_dict()

            for widget in düzenleme_frame_inner.winfo_children():
                widget.destroy()

            def kaydet_ve_kapat():
                try:
                    for kolon, entry in entry_dict.items():
                        yeni_deger = entry.get()
                        if pd.api.types.is_numeric_dtype(df[kolon]):
                            yeni_deger = float(yeni_deger) if '.' in yeni_deger else int(yeni_deger)
                        df.at[satir_index, kolon] = yeni_deger
                    
                    dosya_kaydet()
                    messagebox.showinfo("Başarılı", "Değişiklikler kaydedildi.")
                except ValueError as e:
                    messagebox.showerror("Hata", f"Geçersiz değer: {e}")

            entry_dict = {}
            for idx, (kolon, deger) in enumerate(satir_verileri.items()):
                tk.Label(düzenleme_frame_inner, text=kolon).grid(row=idx, column=0, padx=10, pady=5)
                entry = tk.Entry(düzenleme_frame_inner)
                entry.grid(row=idx, column=1, padx=10, pady=5)
                entry.insert(0, str(deger))
                entry_dict[kolon] = entry

            tk.Button(düzenleme_frame_inner, text="Kaydet", command=kaydet_ve_kapat).grid(row=len(satir_verileri), column=1, pady=10)
        else:
            messagebox.showerror("Hata", "Barkod bulunamadı!")

def dosya_kaydet():
    global dosya_yolu
    if dosya_yolu:
        df.to_excel(dosya_yolu, index=False)
    else:
        messagebox.showerror("Hata", "Dosya yolu bulunamadı.")

def giriş_yap():
    kullanıcı_adı = kullanici_entry.get()
    şifre = sifre_entry.get()
    
    if kullanıcı_adı == "admin" and şifre == "1234":  # Basit doğrulama
        messagebox.showinfo("Başarılı", "Giriş başarılı!")
    else:
        messagebox.showerror("Hata", "Giriş bilgileri yanlış.")


root = tk.Tk()
root.title("NMSExcel")
root.geometry("1366x768")

giris_frame = tk.Frame(root, relief="groove", bd=10)
giris_frame.pack(side="left", padx=20, pady=20, fill="y")

logo = tk.PhotoImage(file="logo.png")
logo_label = tk.Label(giris_frame, image=logo)
logo_label.pack(pady=20)

tk.Label(giris_frame, text="Kullanıcı Adı").pack()
kullanici_entry = tk.Entry(giris_frame)
kullanici_entry.pack(pady=5)

tk.Label(giris_frame, text="Şifre").pack()
sifre_entry = tk.Entry(giris_frame, show="*")
sifre_entry.pack(pady=5)

giris_butonu = GradyanButon(giris_frame, renk1="#4CAF50", renk2="#45A049", text="Giriş Yap", width=15, height=2, command=giriş_yap)
giris_butonu.pack(pady=20)

coded_by_label = tk.Label(giris_frame, text="Coded by NMSHacking", fg="gray")
coded_by_label.pack(pady=10)

secme_frame = tk.Frame(root, relief="groove", bd=10)
secme_frame.pack(side="left", padx=20, pady=20)

excel_butonu = GradyanButon(secme_frame, renk1="#4CAF50", renk2="#45A049", text="Excel Dosyası Seç", command=excel_sec, width=20, height=2)
excel_butonu.pack(pady=20)

barkod_butonu = GradyanButon(secme_frame, renk1="#4CAF50", renk2="#45A049", text="Barkod Gir", command=barkod_no_sor, width=20, height=2)
barkod_butonu.pack(pady=20)

düzenle_butonu = GradyanButon(secme_frame, renk1="#4CAF50", renk2="#45A049", text="Düzenle", command=satir_guncelle, width=20, height=2)
düzenle_butonu.pack(pady=20)
düzenle_butonu.config(state="disabled")

düzenleme_frame_outer = tk.Frame(root, relief="groove", bd=10)
düzenleme_frame_outer.pack(side="left", padx=20, pady=20, fill="both", expand=True)

scrollbar = ttk.Scrollbar(düzenleme_frame_outer)
scrollbar.pack(side="right", fill="y")

canvas = tk.Canvas(düzenleme_frame_outer, yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)

scrollbar.config(command=canvas.yview)

düzenleme_frame_inner = tk.Frame(canvas)
canvas.create_window((0, 0), window=düzenleme_frame_inner, anchor="nw")

düzenleme_frame_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

root.mainloop()
