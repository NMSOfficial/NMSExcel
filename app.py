import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog

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
        satir_guncelle(barkod_no)

def satir_guncelle(barkod):
    global df
    satir_index = df[df['Barkod'] == barkod].index
    if not satir_index.empty:
        satir_index = satir_index[0]
        satir_verileri = df.loc[satir_index].to_dict()

        def kaydet_ve_kapat():
            try:
                for kolon, entry in entry_dict.items():
                    yeni_deger = entry.get()
                    if pd.api.types.is_numeric_dtype(df[kolon]):
                        yeni_deger = float(yeni_deger) if '.' in yeni_deger else int(yeni_deger)
                    df.at[satir_index, kolon] = yeni_deger
                
                dosya_kaydet()
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Hata", f"Geçersiz değer: {e}")

        def dosya_kaydet():
            global dosya_yolu
            if dosya_yolu:
                df.to_excel(dosya_yolu, index=False)
                messagebox.showinfo("Başarılı", "Dosya başarıyla kaydedildi.")
            else:
                messagebox.showerror("Hata", "Dosya yolu bulunamadı.")

        def iptal():
            dialog.destroy()

        dialog = tk.Toplevel(root)
        dialog.title("Veri Güncelle")
        dialog.geometry("500x400")

        
        canvas = tk.Canvas(dialog)
        scrollbar = tk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas)

        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        
        entry_dict = {}
        for idx, (kolon, deger) in enumerate(satir_verileri.items()):
            tk.Label(frame, text=kolon).grid(row=idx, column=0, padx=10, pady=5)
            entry = tk.Entry(frame)
            entry.grid(row=idx, column=1, padx=10, pady=5)
            entry.insert(0, str(deger))
            entry_dict[kolon] = entry

        
        buton_frame = tk.Frame(frame)
        buton_frame.grid(row=len(satir_verileri), columnspan=2, pady=20)

        tk.Button(buton_frame, text="Tamam", command=kaydet_ve_kapat).grid(row=0, column=0, padx=10)
        tk.Button(buton_frame, text="İptal", command=iptal).grid(row=0, column=1, padx=10)

    else:
        messagebox.showerror("Hata", "Barkod bulunamadı!")


root = tk.Tk()
root.title("NMSExcel")
root.geometry("600x600")


logo = tk.PhotoImage(file="logo.png")
logo_label = tk.Label(root, image=logo)
logo_label.pack(pady=20)


excel_butonu = GradyanButon(root, renk1="#4CAF50", renk2="#45A049", text="Excel Seç", command=excel_sec, width=15, height=5)
excel_butonu.pack(pady=20)


root.mainloop()
