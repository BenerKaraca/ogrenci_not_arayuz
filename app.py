import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import pandas as pd
import numpy as np
from tkinter import scrolledtext

class DersBasariHesaplamaArayuz:
    def __init__(self, root):
        self.root = root
        self.root.title("Başarı Hesaplama Sistemi")
        self.root.geometry("800x600")
        
        # Üst menü barı
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # Ders menüsü
        self.ders_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Dersler", menu=self.ders_menu)
        self.ders_menu.add_command(label="Yeni Ders", command=self.yeni_ders)
        self.ders_menu.add_command(label="Ders Listesi", command=self.ders_listesi)
        self.ders_menu.add_separator()
        self.ders_menu.add_command(label="Reset", command=self.reset_form)
        
        # Öğrenci menüsü
        self.ogrenci_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Öğrenciler", menu=self.ogrenci_menu)
        self.ogrenci_menu.add_command(label="Öğrenci Listesi", command=self.ogrenci_listesi)
        
        # Ana frame'i scrollable yapmak için
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack the canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Ana frame'i scrollable frame'e taşı
        self.main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Ders seçim alanı
        self.ders_frame = ttk.LabelFrame(self.main_frame, text="Ders Bilgileri", padding="5")
        self.ders_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Ders kodu ve adı
        ttk.Label(self.ders_frame, text="Ders Kodu:").grid(row=0, column=0, sticky=tk.W)
        self.ders_kodu = ttk.Entry(self.ders_frame, width=15)
        self.ders_kodu.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(self.ders_frame, text="Ders Adı:").grid(row=1, column=0, sticky=tk.W)
        self.ders_adi = ttk.Entry(self.ders_frame, width=30)
        self.ders_adi.grid(row=1, column=1, padx=5, pady=2)
        
        # Excel dosyaları için değişkenler
        self.tablo1_data = None
        self.tablo2_data = None
        self.notlar_data = None
        
        # Excel dosya yükleme frame'i güncelleme
        self.dosya_frame = ttk.LabelFrame(self.main_frame, text="Dosya İşlemleri", padding="5")
        self.dosya_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Tablo 1 yükleme
        self.tablo1_frame = ttk.Frame(self.dosya_frame)
        self.tablo1_frame.grid(row=0, column=0, pady=5)
        ttk.Button(self.tablo1_frame, text="Tablo 1 Yükle", 
                  command=self.yukle_tablo1).grid(row=0, column=0, padx=5)
        self.tablo1_label = ttk.Label(self.tablo1_frame, text="Yüklenmedi", foreground="red")
        self.tablo1_label.grid(row=0, column=1, padx=5)
        ttk.Button(self.tablo1_frame, text="Manuel Giriş", 
                  command=self.manuel_tablo1).grid(row=0, column=2, padx=5)
        
        # Tablo 2 yükleme
        self.tablo2_frame = ttk.Frame(self.dosya_frame)
        self.tablo2_frame.grid(row=1, column=0, pady=5)
        ttk.Button(self.tablo2_frame, text="Tablo 2 Yükle", 
                  command=self.yukle_tablo2).grid(row=0, column=0, padx=5)
        self.tablo2_label = ttk.Label(self.tablo2_frame, text="Yüklenmedi", foreground="red")
        self.tablo2_label.grid(row=0, column=1, padx=5)
        ttk.Button(self.tablo2_frame, text="Manuel Giriş", 
                  command=self.manuel_tablo2).grid(row=0, column=2, padx=5)
        
        # Notlar yükleme
        self.notlar_frame = ttk.Frame(self.dosya_frame)
        self.notlar_frame.grid(row=2, column=0, pady=5)
        ttk.Button(self.notlar_frame, text="Notları Yükle", 
                  command=self.yukle_notlar).grid(row=0, column=0, padx=5)
        self.notlar_label = ttk.Label(self.notlar_frame, text="Yüklenmedi", foreground="red")
        self.notlar_label.grid(row=0, column=1, padx=5)
        ttk.Button(self.notlar_frame, text="Manuel Giriş", 
                  command=self.manuel_notlar).grid(row=0, column=2, padx=5)
        
        # Ders kaydetme ve düzenleme butonları
        self.ders_buton_frame = ttk.Frame(self.ders_frame)
        self.ders_buton_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        ttk.Button(self.ders_buton_frame, text="Ders Kaydet", 
                  command=self.ders_kaydet).grid(row=0, column=0, padx=5)

        # Tablo işlemleri için butonlar güncelleme
        # Tablo 1
        self.tablo1_buton_frame = ttk.Frame(self.tablo1_frame)
        self.tablo1_buton_frame.grid(row=1, column=0, columnspan=3, pady=5)
        ttk.Button(self.tablo1_buton_frame, text="Tabloyu İndir", 
                  command=lambda: self.excel_kaydet(self.tablo1_data, "Tablo1.xlsx")).grid(row=0, column=0, padx=5)
        ttk.Button(self.tablo1_buton_frame, text="Temizle", 
                  command=lambda: self.tablo_temizle("tablo1")).grid(row=0, column=1, padx=5)

        # Tablo 2
        self.tablo2_buton_frame = ttk.Frame(self.tablo2_frame)
        self.tablo2_buton_frame.grid(row=1, column=0, columnspan=3, pady=5)
        ttk.Button(self.tablo2_buton_frame, text="Tabloyu İndir", 
                  command=lambda: self.excel_kaydet(self.tablo2_data, "Tablo2.xlsx")).grid(row=0, column=0, padx=5)
        ttk.Button(self.tablo2_buton_frame, text="Temizle", 
                  command=lambda: self.tablo_temizle("tablo2")).grid(row=0, column=1, padx=5)

        # Notlar
        self.notlar_buton_frame = ttk.Frame(self.notlar_frame)
        self.notlar_buton_frame.grid(row=1, column=0, columnspan=3, pady=5)
        ttk.Button(self.notlar_buton_frame, text="Tabloyu İndir", 
                  command=lambda: self.excel_kaydet(self.notlar_data, "Notlar.xlsx")).grid(row=0, column=0, padx=5)
        ttk.Button(self.notlar_buton_frame, text="Temizle", 
                  command=lambda: self.tablo_temizle("notlar")).grid(row=0, column=1, padx=5)

        # Değerlendirme kriterleri için değişkenler
        self.kriterler = []
        self.kriter_agirliklari = []
        
        # Değerlendirme kriterleri frame'i
        self.kriter_frame = ttk.LabelFrame(self.main_frame, text="Değerlendirme Kriterleri", padding="5")
        self.kriter_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Kriter ekleme alanı
        self.kriter_input_frame = ttk.Frame(self.kriter_frame)
        self.kriter_input_frame.grid(row=0, column=0, columnspan=2, pady=5)
        
        ttk.Label(self.kriter_input_frame, text="Kriter:").grid(row=0, column=0, padx=5)
        self.kriter_adi = ttk.Entry(self.kriter_input_frame, width=20)
        self.kriter_adi.grid(row=0, column=1, padx=5)
        
        ttk.Label(self.kriter_input_frame, text="Ağırlık (%):").grid(row=0, column=2, padx=5)
        self.kriter_agirlik = ttk.Entry(self.kriter_input_frame, width=10)
        self.kriter_agirlik.grid(row=0, column=3, padx=5)
        
        ttk.Button(self.kriter_input_frame, text="Kriter Ekle", 
                  command=self.kriter_ekle).grid(row=0, column=4, padx=5)
        
        # Kriter listesi görüntüleme alanı
        self.kriter_liste_frame = ttk.Frame(self.kriter_frame)
        self.kriter_liste_frame.grid(row=1, column=0, columnspan=2, pady=5)
        
        self.kriter_liste = scrolledtext.ScrolledText(self.kriter_liste_frame, 
                                                    width=40, height=8, wrap=tk.WORD)
        self.kriter_liste.grid(row=0, column=0, padx=5)
        
        # Toplam ağırlık göstergesi
        self.toplam_agirlik_label = ttk.Label(self.kriter_frame, text="Toplam Ağırlık: 0%")
        self.toplam_agirlik_label.grid(row=2, column=0, pady=5)
        
        # Kaydet ve Düzenle butonları
        self.buton_frame = ttk.Frame(self.kriter_frame)
        self.buton_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        ttk.Button(self.buton_frame, text="Kriterleri Kaydet", 
                  command=self.kriterleri_kaydet).grid(row=0, column=0, padx=5)
        ttk.Button(self.buton_frame, text="Kriterleri Düzenle", 
                  command=self.kriterleri_duzenle).grid(row=0, column=1, padx=5)

        # Ders listesi için değişkenler
        self.dersler = []  # (kod, ad) tuple'ları
        self.aktif_ders_index = None
        
        # Ana frame'e Reset butonu ekleme
        self.reset_btn = ttk.Button(self.main_frame, text="Reset - Ders Seçimine Dön", 
                                  style='Action.TButton', command=self.reset_form)
        self.reset_btn.grid(row=4, column=0, pady=10)

        # Sonuç tabloları frame'i
        self.sonuc_frame = ttk.LabelFrame(self.main_frame, text="Sonuç Tabloları", padding="5")
        self.sonuc_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Tablo 4 ve 5 oluşturma butonları
        ttk.Button(self.sonuc_frame, text="Tablo 4 Olustur", 
                  command=self.tablo4_olustur).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(self.sonuc_frame, text="Tablo 5 Olustur", 
                  command=self.tablo5_olustur).grid(row=0, column=1, padx=5, pady=5)

    def yeni_ders(self):
        """Yeni ders ekleme"""
        if len(self.dersler) >= 3 and not messagebox.askyesno(
            "Uyarı", "Zaten 3 ders eklenmiş. Yeni ders eklemek istiyor musunuz?"):
            return
            
        self.reset_form()
        self.aktif_ders_index = None

    def ders_listesi(self):
        """Ders listesi penceresi"""
        liste_pencere = tk.Toplevel(self.root)
        liste_pencere.title("Ders Listesi")
        liste_pencere.geometry("400x300")
        
        # Ders listesi
        liste_frame = ttk.Frame(liste_pencere, padding="10")
        liste_frame.pack(fill=tk.BOTH, expand=True)
        
        # Başlık
        ttk.Label(liste_frame, text="Kayıtlı Dersler", 
                 style='Header.TLabel').pack(pady=(0,10))
        
        # Dersler
        for i, (kod, ad) in enumerate(self.dersler):
            ders_frame = ttk.Frame(liste_frame)
            ders_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(ders_frame, text=f"{kod} - {ad}").pack(side=tk.LEFT)
            ttk.Button(ders_frame, text="Seç", 
                      command=lambda idx=i: self.ders_sec(idx)).pack(side=tk.RIGHT)

    def ders_sec(self, index):
        """Seçilen dersi yükle"""
        self.aktif_ders_index = index
        kod, ad = self.dersler[index]
        
        # Form alanlarını doldur
        self.ders_kodu.delete(0, tk.END)
        self.ders_kodu.insert(0, kod)
        
        self.ders_adi.delete(0, tk.END)
        self.ders_adi.insert(0, ad)
        
        # TODO: Diğer ders verilerini yükle

    def ogrenci_listesi(self):
        """Öğrenci listesi penceresi"""
        liste_pencere = tk.Toplevel(self.root)
        liste_pencere.title("Öğrenci Listesi")
        liste_pencere.geometry("600x400")
        
        # Excel yükleme
        ttk.Button(liste_pencere, text="Excel'den Öğrenci Listesi Yükle",
                  command=self.yukle_ogrenci_listesi).pack(pady=10)
        
        # Manuel öğrenci ekleme
        ekle_frame = ttk.Frame(liste_pencere)
        ekle_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(ekle_frame, text="Öğrenci No:").pack(side=tk.LEFT)
        ogrenci_no = ttk.Entry(ekle_frame, width=15)
        ogrenci_no.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(ekle_frame, text="Ekle",
                  command=lambda: self.ogrenci_ekle(ogrenci_no.get())).pack(side=tk.LEFT)

    def yukle_ogrenci_listesi(self):
        """Excel'den öğrenci listesi yükleme"""
        filename = filedialog.askopenfilename(
            title="Öğrenci Listesi Seç",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                df = pd.read_excel(filename)
                if 'Öğrenci No' not in df.columns:
                    raise ValueError("Excel dosyasında 'Öğrenci No' sütunu bulunamadı!")
                # TODO: Öğrenci listesini kaydet
                messagebox.showinfo("Başarılı", "Öğrenci listesi yüklendi!")
            except Exception as e:
                messagebox.showerror("Hata", str(e))

    def durum_etiketi_guncelle(self, etiket, durum):
        """Durum etiketini güncelle"""
        if durum:
            etiket.config(text="Yüklendi ✓", foreground="green")
        else:
            etiket.config(text="Yüklenmedi ✗", foreground="red")

    def reset_form(self):
        """Formu temizle ve ders seçimine dön"""
        # Form alanlarını temizle
        self.ders_kodu.delete(0, tk.END)
        self.ders_adi.delete(0, tk.END)
        self.kriter_liste.delete(1.0, tk.END)
        self.toplam_agirlik_label.config(text="Toplam Ağırlık: 0%")
        
        # Durum etiketlerini sıfırla
        for label in [self.tablo1_label, self.tablo2_label, self.notlar_label]:
            self.durum_etiketi_guncelle(label, False)
        
        # Tüm verileri temizle
        self.tablo1_data = None
        self.tablo2_data = None
        self.notlar_data = None
        self.kriterler = []
        self.kriter_agirliklari = []
        
        # Aktif dersi sıfırla
        self.aktif_ders_index = None
        
        messagebox.showinfo("Bilgi", "Form temizlendi, ders seçimine dönüldü.")

    def yukle_tablo1(self):
        filename = filedialog.askopenfilename(
            title="Tablo 1'i Seç",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                df = pd.read_excel(filename)
                # Veri doğrulama
                if self.dogrula_tablo1(df):
                    self.tablo1_data = df
                    self.tablo1_label.config(text="Yüklendi ✓", foreground="green")
                    messagebox.showinfo("Başarılı", "Tablo 1 başarıyla yüklendi!")
            except Exception as e:
                messagebox.showerror("Hata", f"Dosya yüklenirken hata oluştu: {str(e)}")

    def yukle_tablo2(self):
        filename = filedialog.askopenfilename(
            title="Tablo 2'yi Seç",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                df = pd.read_excel(filename)
                # Veri doğrulama
                if self.dogrula_tablo2(df):
                    self.tablo2_data = df
                    self.tablo2_label.config(text="Yüklendi ✓", foreground="green")
                    messagebox.showinfo("Başarılı", "Tablo 2 başarıyla yüklendi!")
            except Exception as e:
                messagebox.showerror("Hata", f"Dosya yüklenirken hata oluştu: {str(e)}")

    def yukle_notlar(self):
        """Notları Excel'den yükle"""
        filename = filedialog.askopenfilename(
            title="Not Listesi Seç",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                df = pd.read_excel(filename)
                
                # Öğrenci numarası sütunu hariç diğer sütunları kontrol et
                numeric_columns = df.select_dtypes(include=[np.number]).columns
                for col in numeric_columns:
                    if col != 'Ogrenci_No':  # Öğrenci numarası sütununu atla
                        if not df[col].between(0, 100).all():
                            raise ValueError(f"{col} sütunundaki tüm notlar 0-100 arasında olmalıdır!")
                
                self.notlar_data = df
                self.notlar_label.config(text="Yüklendi ✓", foreground="green")
                messagebox.showinfo("Başarılı", "Notlar başarıyla yüklendi!")
                
            except Exception as e:
                messagebox.showerror("Hata", str(e))
                self.notlar_label.config(text="Yüklenmedi", foreground="red")

    def dogrula_tablo1(self, df):
        try:
            # İlk sütunu (Prg Çıktı) hariç tutarak sayısal sütunları kontrol et
            numeric_columns = df.select_dtypes(include=[np.number]).columns[1:]  # İlk sütunu atla
            
            if not ((df[numeric_columns] >= 0) & (df[numeric_columns] <= 1)).all().all():
                messagebox.showerror("Hata", "Tablo 1'deki değerler 0 ile 1 arasında olmalıdır!")
                return False
            return True
        
        except Exception as e:
            messagebox.showerror("Hata", f"Tablo 1 doğrulama hatası: {str(e)}")
            return False

    def dogrula_tablo2(self, df):
        try:
            # Değerlerin [0,1] aralığında olduğunu kontrol et
            if not ((df.select_dtypes(include=[np.number]) >= 0) & 
                   (df.select_dtypes(include=[np.number]) <= 1)).all().all():
                messagebox.showerror("Hata", "Tablo 2'deki değerler 0 ile 1 arasında olmalıdır!")
                return False
            return True
        except Exception as e:
            messagebox.showerror("Hata", f"Tablo 2 doğrulama hatası: {str(e)}")
            return False

    def dogrula_notlar(self, df):
        try:
            # Öğrenci numarası sütununun varlığını kontrol et
            if 'Öğrenci No' not in df.columns:
                messagebox.showerror("Hata", "Notlar tablosunda 'Öğrenci No' sütunu bulunmalıdır!")
                return False
            return True
        except Exception as e:
            messagebox.showerror("Hata", f"Notlar doğrulama hatası: {str(e)}")
            return False

    def manuel_tablo1(self):
        self.manuel_pencere = tk.Toplevel(self.root)
        self.manuel_pencere.title("Tablo 1 Manuel Giriş")
        self.manuel_pencere.geometry("800x600")

        # Tablo boyutları için giriş
        boyut_frame = ttk.LabelFrame(self.manuel_pencere, text="Tablo Boyutları", padding="5")
        boyut_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(boyut_frame, text="Program Çıktı Sayısı:").grid(row=0, column=0, padx=5)
        self.prog_cikti_sayisi = ttk.Entry(boyut_frame, width=10)
        self.prog_cikti_sayisi.grid(row=0, column=1, padx=5)

        ttk.Label(boyut_frame, text="Ders Çıktı Sayısı:").grid(row=0, column=2, padx=5)
        self.ders_cikti_sayisi = ttk.Entry(boyut_frame, width=10)
        self.ders_cikti_sayisi.grid(row=0, column=3, padx=5)

        ttk.Button(boyut_frame, text="Tablo Oluştur", 
                  command=self.tablo1_olustur).grid(row=0, column=4, padx=5)

    def tablo1_olustur(self):
        try:
            prog_sayisi = int(self.prog_cikti_sayisi.get())
            ders_sayisi = int(self.ders_cikti_sayisi.get())

            # Tablo frame
            self.tablo1_frame = ttk.Frame(self.manuel_pencere)
            self.tablo1_frame.pack(fill="both", expand=True, padx=5, pady=5)

            # Başlık satırı
            ttk.Label(self.tablo1_frame, text="Program/Ders").grid(row=0, column=0, padx=2, pady=2)
            for j in range(ders_sayisi):
                ttk.Label(self.tablo1_frame, text=f"Ders Çıktı {j+1}").grid(row=0, column=j+1, padx=2, pady=2)

            # Tablo hücreleri
            self.tablo1_cells = []
            for i in range(prog_sayisi):
                row_cells = []
                ttk.Label(self.tablo1_frame, text=f"Program Çıktı {i+1}").grid(row=i+1, column=0, padx=2, pady=2)
                for j in range(ders_sayisi):
                    cell = ttk.Entry(self.tablo1_frame, width=8)
                    cell.grid(row=i+1, column=j+1, padx=2, pady=2)
                    cell.insert(0, "0")
                    row_cells.append(cell)
                self.tablo1_cells.append(row_cells)

            # Kaydet butonu
            ttk.Button(self.manuel_pencere, text="Kaydet", 
                      command=self.tablo1_kaydet).pack(pady=10)

        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli sayılar girin!")

    def tablo1_kaydet(self):
        try:
            data = []
            for row in self.tablo1_cells:
                row_data = []
                for cell in row:
                    value = float(cell.get())
                    if not (0 <= value <= 1):
                        raise ValueError("Değerler 0 ile 1 arasında olmalıdır!")
                    row_data.append(value)
                data.append(row_data)

            self.tablo1_data = pd.DataFrame(data)
            self.tablo1_label.config(text="Yüklendi ✓", foreground="green")
            self.manuel_pencere.destroy()
            messagebox.showinfo("Başarılı", "Tablo 1 başarıyla kaydedildi!")

        except ValueError as e:
            messagebox.showerror("Hata", str(e))

    def manuel_tablo2(self):
        self.manuel_pencere = tk.Toplevel(self.root)
        self.manuel_pencere.title("Tablo 2 Manuel Giriş")
        self.manuel_pencere.geometry("800x600")

        # Tablo boyutları için giriş
        boyut_frame = ttk.LabelFrame(self.manuel_pencere, text="Tablo Boyutları", padding="5")
        boyut_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(boyut_frame, text="Ders Çıktı Sayısı:").grid(row=0, column=0, padx=5)
        self.ders_cikti_sayisi = ttk.Entry(boyut_frame, width=10)
        self.ders_cikti_sayisi.grid(row=0, column=1, padx=5)

        ttk.Label(boyut_frame, text="Değerlendirme Kriteri Sayısı:").grid(row=0, column=2, padx=5)
        self.kriter_sayisi = ttk.Entry(boyut_frame, width=10)
        self.kriter_sayisi.grid(row=0, column=3, padx=5)

        ttk.Button(boyut_frame, text="Tablo Oluştur", 
                  command=self.tablo2_olustur).grid(row=0, column=4, padx=5)

    def tablo2_olustur(self):
        try:
            ders_sayisi = int(self.ders_cikti_sayisi.get())
            kriter_sayisi = int(self.kriter_sayisi.get())

            # Tablo frame
            self.tablo2_frame = ttk.Frame(self.manuel_pencere)
            self.tablo2_frame.pack(fill="both", expand=True, padx=5, pady=5)

            # Başlık satırı
            ttk.Label(self.tablo2_frame, text="Ders/Kriter").grid(row=0, column=0, padx=2, pady=2)
            for j in range(kriter_sayisi):
                ttk.Label(self.tablo2_frame, text=f"Kriter {j+1}").grid(row=0, column=j+1, padx=2, pady=2)

            # Tablo hücreleri
            self.tablo2_cells = []
            for i in range(ders_sayisi):
                row_cells = []
                ttk.Label(self.tablo2_frame, text=f"Ders Çıktı {i+1}").grid(row=i+1, column=0, padx=2, pady=2)
                for j in range(kriter_sayisi):
                    cell = ttk.Entry(self.tablo2_frame, width=8)
                    cell.grid(row=i+1, column=j+1, padx=2, pady=2)
                    cell.insert(0, "0")
                    row_cells.append(cell)
                self.tablo2_cells.append(row_cells)

            # Kaydet butonu
            ttk.Button(self.manuel_pencere, text="Kaydet", 
                      command=self.tablo2_kaydet).pack(pady=10)

        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli sayılar girin!")

    def tablo2_kaydet(self):
        try:
            data = []
            for row in self.tablo2_cells:
                row_data = []
                for cell in row:
                    value = float(cell.get())
                    if not (0 <= value <= 1):
                        raise ValueError("Değerler 0 ile 1 arasında olmalıdır!")
                    row_data.append(value)
                data.append(row_data)

            self.tablo2_data = pd.DataFrame(data)
            self.tablo2_label.config(text="Yüklendi ✓", foreground="green")
            self.manuel_pencere.destroy()
            messagebox.showinfo("Başarılı", "Tablo 2 başarıyla kaydedildi!")

        except ValueError as e:
            messagebox.showerror("Hata", str(e))

    def manuel_notlar(self):
        """Notları manuel giriş penceresi"""
        self.manuel_pencere = tk.Toplevel(self.root)
        self.manuel_pencere.title("Notlar Manuel Giriş")
        self.manuel_pencere.geometry("800x400")

        # Öğrenci sayısı için giriş
        ogrenci_frame = ttk.LabelFrame(self.manuel_pencere, text="Öğrenci Bilgileri", padding="5")
        ogrenci_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(ogrenci_frame, text="Öğrenci Sayısı:").grid(row=0, column=0, padx=5)
        self.ogrenci_sayisi = ttk.Entry(ogrenci_frame, width=10)
        self.ogrenci_sayisi.grid(row=0, column=1, padx=5)

        ttk.Button(ogrenci_frame, text="Tablo Oluştur", 
                  command=self.notlar_tablo_olustur).grid(row=0, column=2, padx=5)

    def notlar_tablo_olustur(self):
        try:
            ogrenci_sayisi = int(self.ogrenci_sayisi.get())

            # Tablo frame
            self.notlar_frame = ttk.Frame(self.manuel_pencere)
            self.notlar_frame.pack(fill="both", expand=True, padx=5, pady=5)

            # Başlıklar - Sabit sütun isimleri
            columns = ['Ogrenci_No', 'Odev1', 'Odev2', 'Quiz', 'Vize', 'Final']
            for i, col in enumerate(columns):
                ttk.Label(self.notlar_frame, text=col).grid(row=0, column=i, padx=5, pady=2)

            # Öğrenci notları için giriş alanları
            self.not_cells = []
            for i in range(ogrenci_sayisi):
                row_cells = []
                for j, col in enumerate(columns):
                    cell = ttk.Entry(self.notlar_frame, width=10)
                    cell.grid(row=i+1, column=j, padx=5, pady=2)
                    if j > 0:  # Öğrenci No hariç diğer hücrelere 0 değeri
                        cell.insert(0, "0")
                    row_cells.append(cell)
                self.not_cells.append(row_cells)

            # Kaydet butonu
            ttk.Button(self.manuel_pencere, text="Kaydet", 
                      command=self.notlar_kaydet).pack(pady=10)

        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli bir öğrenci sayısı girin!")

    def notlar_kaydet(self):
        try:
            data = []
            columns = ['Ogrenci_No', 'Odev1', 'Odev2', 'Quiz', 'Vize', 'Final']
            
            for row in self.not_cells:
                row_data = {}
                for i, cell in enumerate(row):
                    value = cell.get().strip()
                    if i == 0:  # Öğrenci No kontrolü
                        if not value:
                            raise ValueError("Öğrenci numarası boş bırakılamaz!")
                        row_data[columns[i]] = value
                    else:  # Not kontrolü
                        value = float(value)
                        if not (0 <= value <= 100):
                            raise ValueError(f"{columns[i]} için notlar 0-100 arasında olmalıdır!")
                        row_data[columns[i]] = value
                data.append(row_data)

            self.notlar_data = pd.DataFrame(data)
            self.notlar_label.config(text="Yüklendi ✓", foreground="green")
            self.manuel_pencere.destroy()
            messagebox.showinfo("Başarılı", "Notlar başarıyla kaydedildi!")

        except ValueError as e:
            messagebox.showerror("Hata", str(e))

    def kriter_ekle(self):
        """Kriter ekleme fonksiyonu"""
        kriter = self.kriter_adi.get().strip()
        try:
            agirlik = float(self.kriter_agirlik.get())
            if agirlik <= 0 or agirlik > 100:
                raise ValueError("Ağırlık 0-100 arasında olmalıdır!")
                
            # Minimum 5 kriter kontrolü eklenecek
            if len(self.kriterler) >= 5 and not messagebox.askyesno(
                "Uyarı", "Minimum 5 kriter girildi. Devam etmek istiyor musunuz?"):
                return
                
            self.kriterler.append((kriter, agirlik))
            self.kriter_listesini_guncelle()
            
            # Form temizleme
            self.kriter_adi.delete(0, tk.END)
            self.kriter_agirlik.delete(0, tk.END)
            
        except ValueError as e:
            messagebox.showerror("Hata", str(e))

    def kriter_listesini_guncelle(self):
        self.kriter_liste.delete(1.0, tk.END)
        toplam_agirlik = 0
        
        for i, (kriter, agirlik) in enumerate(self.kriterler, 1):
            self.kriter_liste.insert(tk.END, 
                                   f"{i}. {kriter}: %{agirlik}\n")
            toplam_agirlik += float(agirlik)
        
        self.toplam_agirlik_label.config(text=f"Toplam Ağırlık: %{toplam_agirlik}")
        
        # Minimum 5 kriter kontrolü
        if len(self.kriterler) >= 5 and abs(toplam_agirlik - 100) < 0.01:
            self.kriter_liste.insert(tk.END, "\n✓ Kriterler geçerli!")
        else:
            self.kriter_liste.insert(tk.END, 
                "\n⚠ En az 5 kriter gerekli ve toplam ağırlık 100 olmalıdır!")

    def kriterleri_kaydet(self):
        """Kriterleri ve ders verilerini kaydet"""
        if not self.kriterler:
            messagebox.showerror("Hata", "Kaydedilecek kriter bulunmamaktadır!")
            return
        
        # Minimum 5 kriter kontrolü
        if len(self.kriterler) < 5:
            messagebox.showerror("Hata", "En az 5 kriter girilmelidir!")
            return
        
        # Toplam ağırlık kontrolü
        toplam_agirlik = sum(agirlik for _, agirlik in self.kriterler)
        if abs(toplam_agirlik - 100) > 0.01:
            messagebox.showerror("Hata", "Toplam ağırlık 100 olmalıdır!")
            return
        
        # Aktif ders varsa verilerini kaydet
        if self.aktif_ders_index is not None:
            self.dersler[self.aktif_ders_index] = (
                self.ders_kodu.get(),
                self.ders_adi.get()
            )
        else:
            self.dersler.append((
                self.ders_kodu.get(),
                self.ders_adi.get()
            ))
            self.aktif_ders_index = len(self.dersler) - 1
        
        messagebox.showinfo("Başarılı", "Kriterler ve ders verileri kaydedildi!")

    def kriterleri_duzenle(self):
        if not self.kriterler:
            messagebox.showinfo("Bilgi", "Düzenlenecek kriter bulunmamaktadır!")
            return
            
        # Düzenleme penceresi oluştur
        duzen_pencere = tk.Toplevel(self.root)
        duzen_pencere.title("Kriterleri Düzenle")
        duzen_pencere.geometry("400x300")
        
        # Kriter listesi ve düzenleme alanı
        for i, (kriter, agirlik) in enumerate(self.kriterler):
            frame = ttk.Frame(duzen_pencere)
            frame.pack(pady=5, padx=10, fill=tk.X)
            
            kriter_entry = ttk.Entry(frame, width=20)
            kriter_entry.insert(0, kriter)
            kriter_entry.pack(side=tk.LEFT, padx=5)
            
            agirlik_entry = ttk.Entry(frame, width=10)
            agirlik_entry.insert(0, agirlik)
            agirlik_entry.pack(side=tk.LEFT, padx=5)
            
            # Silme butonu
            ttk.Button(frame, text="Sil", 
                      command=lambda idx=i: self.kriter_sil(idx, duzen_pencere)).pack(side=tk.LEFT, padx=5)

    def kriter_sil(self, index, pencere):
        del self.kriterler[index]
        self.kriter_listesini_guncelle()
        pencere.destroy()
        self.kriterleri_duzenle()

    def hesapla_ve_kaydet(self):
        """Hesaplama ve kaydetme işlemi sırasında ilerleme göstergesi"""
        if not self.kontrol_veriler():
            return
        
        try:
            # İlerleme penceresi
            progress_window = tk.Toplevel(self.root)
            progress_window.title("İşlem Durumu")
            progress_window.geometry("300x150")
            
            progress_label = ttk.Label(progress_window, text="İşlem yapılıyor...")
            progress_label.pack(pady=20)
            
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
            progress_bar.pack(pady=10, padx=20, fill=tk.X)
            progress_bar.start()
            
            # Hesaplama işlemleri
            self.hesapla_tablo3()
            progress_label.config(text="Tablo 3 hesaplandı...")
            
            self.hesapla_tablo4()
            progress_label.config(text="Tablo 4 hesaplandı...")
            
            self.hesapla_tablo5()
            progress_label.config(text="Tablo 5 hesaplandı...")
            
            self.sonuclari_kaydet()
            progress_label.config(text="Sonuçlar kaydedildi!")
            
            progress_bar.stop()
            progress_window.destroy()
            
            messagebox.showinfo("Başarılı", "Hesaplamalar tamamlandı ve sonuçlar kaydedildi!")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama sırasında hata oluştu: {str(e)}")

    def tablo_butonlari_olustur(self, frame, kaydet_komut, temizle_komut):
        """Tablo işlemleri için butonları oluştur"""
        buton_frame = ttk.Frame(frame)
        buton_frame.grid(row=1, column=0, pady=5)
        
        ttk.Button(buton_frame, text="Kaydet", 
                  command=kaydet_komut).grid(row=0, column=0, padx=5)
        ttk.Button(buton_frame, text="Temizle", 
                  command=temizle_komut).grid(row=0, column=1, padx=5)

    def excel_kaydet(self, data, dosya_adi):
        """Excel dosyasını kaydet"""
        try:
            kayit_yolu = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=dosya_adi
            )
            if kayit_yolu:
                data.to_excel(kayit_yolu, index=False)
                messagebox.showinfo("Başarılı", f"Dosya kaydedildi: {kayit_yolu}")
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya kaydedilirken hata oluştu: {str(e)}")

    def sonuc_butonlari_olustur(self):
        """Tablo 4 ve 5 oluşturma butonları"""
        sonuc_frame = ttk.LabelFrame(self.main_frame, text="Sonuç Tabloları", padding="5")
        sonuc_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Button(sonuc_frame, text="Tablo 4 Olustur", 
                  command=self.tablo4_olustur).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(sonuc_frame, text="Tablo 5 Olustur", 
                  command=self.tablo5_olustur).grid(row=0, column=1, padx=5, pady=5)

    def verileri_dogrula(self):
        """Tüm gerekli verilerin yüklenip yüklenmediğini kontrol et"""
        if self.tablo1_data is None or self.tablo2_data is None or self.notlar_data is None:
            messagebox.showerror("Hata", "Lütfen önce tüm tabloları yükleyin!")
            return False
        return True

    def tablo4_olustur(self):
        """Tablo 4 oluştur"""
        if not self.verileri_dogrula():
            return
        
        try:
            # Ders bilgilerini kontrol et
            if not self.ders_kodu.get() or not self.ders_adi.get():
                messagebox.showerror("Hata", "Lütfen ders kodu ve adını girin!")
                return
            
            # Kriter ağırlıklarını kontrol et
            if not self.kriterler:
                messagebox.showerror("Hata", "Lütfen değerlendirme kriterlerini ekleyin!")
                return
            
            # Kriter ağırlıklarını al
            oranlar = [agirlik for _, agirlik in self.kriterler]
            
            # Her öğrenci için ders çıktıları başarı oranlarını hesapla
            ders_ciktilari_basari_oranlari = []
            
            # DataFrame'lerden değerleri doğru şekilde al
            ogrenci_notlari = self.notlar_data.iloc[:, 1:6].values  # Sadece not sütunlarını al
            ogrenci_nolar = self.notlar_data['Ogrenci_No'].values
            iliski_matrisi = self.tablo2_data.iloc[1:7, 1:6].values  # Başlık satırını atlayarak al
            
            # Her öğrenci için hesaplama yap
            for ogrenci in ogrenci_notlari:
                basari_oranlari = []
                
                # Her ders çıktısı için hesaplama yap
                for i in range(len(iliski_matrisi)):
                    toplam = 0
                    for j in range(len(iliski_matrisi[i])):
                        # İlişki matrisi (0/1) * öğrenci notu * değerlendirme kriteri yüzdesi
                        deger = iliski_matrisi[i][j] * (ogrenci[j]/10) * (oranlar[j] / 100)
                        toplam += deger

                    # Max değeri hesaplama: İlişki matrisi satırındaki 1'lerin toplam ağırlığı * 100
                    max_not = sum(iliski_matrisi[i][j] * (oranlar[j] / 100) 
                                for j in range(len(iliski_matrisi[i]))) * 10

                    # Başarı yüzdesi hesaplama
                    yuzde_basari = (toplam / max_not) * 100 if max_not != 0 else 0

                    basari_oranlari.append({
                        'Ders Çıktı': f'Ders Çıktı {i+1}',
                        'Ödev1': (ogrenci[0]/10) if iliski_matrisi[i][0] == 1 else 0,
                        'Ödev2': (ogrenci[1]/10) if iliski_matrisi[i][1] == 1 else 0,
                        'Quiz': (ogrenci[2]/10) if iliski_matrisi[i][2] == 1 else 0,
                        'Vize': (ogrenci[3]/10) if iliski_matrisi[i][3] == 1 else 0,
                        'Final': (ogrenci[4]/10) if iliski_matrisi[i][4] == 1 else 0,
                        'Toplam': toplam,
                        'Max': max_not,
                        '% Başarı': yuzde_basari
                    })
                ders_ciktilari_basari_oranlari.append(basari_oranlari)

            # Excel'e kaydet
            kayit_yolu = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f'{self.ders_kodu.get()}.Tablo4.xlsx'
            )
            
            if kayit_yolu:
                with pd.ExcelWriter(kayit_yolu) as writer:
                    for idx, ogrenci_no in enumerate(ogrenci_nolar):
                        df = pd.DataFrame(ders_ciktilari_basari_oranlari[idx])
                        df.to_excel(writer, sheet_name=f'Öğrenci {ogrenci_no}', index=False)
            
            messagebox.showinfo("Başarılı", "Tablo 4 oluşturuldu!")
        
        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama sırasında hata oluştu: {str(e)}")

    def tablo5_olustur(self):
        """Tablo 5 oluştur"""
        if not self.verileri_dogrula():
            return
        
        try:
            # Ders bilgilerini kontrol et
            if not self.ders_kodu.get() or not self.ders_adi.get():
                messagebox.showerror("Hata", "Lütfen ders kodu ve adını girin!")
                return
            
            # Kriter ağırlıklarını kontrol et
            if not self.kriterler:
                messagebox.showerror("Hata", "Lütfen değerlendirme kriterlerini ekleyin!")
                return
            
            # Tablo 1'den program çıktıları ilişki matrisini al ve yuvarla
            program_ciktilari_iliski = self.tablo1_data.iloc[0:10, 1:6].values
            program_ciktilari_iliski = np.round(program_ciktilari_iliski).astype(int)
            
            # Tablo 2'den ders çıktıları ilişki matrisini al
            iliski_matrisi = self.tablo2_data.iloc[1:7, 1:6].values
            
            # Kriter ağırlıklarını al
            oranlar = [agirlik for _, agirlik in self.kriterler]
            
            # Öğrenci notlarını al
            ogrenci_notlari = self.notlar_data.iloc[:, 1:].values
            ogrenci_nolar = self.notlar_data['Ogrenci_No'].values
            
            # Önce Tablo 4'ün verilerini hesapla
            ders_ciktilari_basari_oranlari = []
            for ogrenci in ogrenci_notlari:
                basari_oranlari = []
                for i in range(len(iliski_matrisi)):
                    toplam = 0
                    for j in range(len(iliski_matrisi[i])):
                        deger = iliski_matrisi[i][j] * (ogrenci[j]/10) * (oranlar[j] / 100)
                        toplam += deger

                    max_not = sum(iliski_matrisi[i][j] * (oranlar[j] / 100) 
                                for j in range(len(iliski_matrisi[i]))) * 10
                    yuzde_basari = (toplam / max_not) * 100 if max_not != 0 else 0
                    basari_oranlari.append({'% Başarı': yuzde_basari})
                ders_ciktilari_basari_oranlari.append(basari_oranlari)

            # Program çıktıları başarı oranlarını hesapla
            program_ciktilari_basari_oranlari = []
            for basari_oranlari in ders_ciktilari_basari_oranlari:
                program_basari_oranlari = []
                
                for i in range(len(program_ciktilari_iliski)):
                    ders_basarilari = []
                    toplam_iliski = 0
                    toplam_basari = 0
                    
                    for j in range(5):
                        iliski_degeri = program_ciktilari_iliski[i][j]
                        if iliski_degeri > 0:
                            toplam_iliski += iliski_degeri
                            toplam_basari += iliski_degeri * basari_oranlari[j]['% Başarı']
                        ders_basarilari.append(
                            basari_oranlari[j]['% Başarı'] if iliski_degeri > 0 else 0.0
                        )
                    
                    basari_orani = round(toplam_basari / toplam_iliski, 1) if toplam_iliski > 0 else 0.0
                    program_basari_oranlari.append({
                        'Program Çıktı': i + 1,
                        'Ders Çıktıları': ders_basarilari,
                        'Başarı Oranı': basari_orani
                    })
                program_ciktilari_basari_oranlari.append(program_basari_oranlari)

            # Excel'e kaydet
            with pd.ExcelWriter(f'Tablo5_{self.ders_kodu.get()}.xlsx') as writer:
                # Ders bilgilerini ekle
                ders_bilgileri = pd.DataFrame({
                    'Ders Kodu': [self.ders_kodu.get()],
                    'Ders Adı': [self.ders_adi.get()]
                })
                ders_bilgileri.to_excel(writer, sheet_name='Ders Bilgileri', index=False)
                
                # Kriter bilgilerini ekle
                kriter_bilgileri = pd.DataFrame(self.kriterler, columns=['Kriter', 'Ağırlık'])
                kriter_bilgileri.to_excel(writer, sheet_name='Değerlendirme Kriterleri', index=False)
                
                # Her öğrenci için sonuçları ekle
                for idx, ogrenci_no in enumerate(ogrenci_nolar):
                    basari_data = []
                    for prog in program_ciktilari_basari_oranlari[idx]:
                        row = [prog['Program Çıktı']] + prog['Ders Çıktıları'] + [prog['Başarı Oranı']]
                        basari_data.append(row)
                    
                    columns = ['Prg Çıktı'] + [f'Ders çıktısı {i}' for i in range(1, 6)] + ['Başarı Oranı']
                    df = pd.DataFrame(basari_data, columns=columns)
                    df.to_excel(writer, sheet_name=f'Öğrenci {ogrenci_no}', index=False)
            
            messagebox.showinfo("Başarılı", "Tablo 5 oluşturuldu!")
            
        except Exception as e:
            messagebox.showerror("Hata", str(e))

    def ders_butonlari_olustur(self):
        """Ders işlemleri butonları"""
        ders_buton_frame = ttk.Frame(self.ders_frame)
        ders_buton_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        ttk.Button(ders_buton_frame, text="Ders Kaydet", 
                  command=self.ders_kaydet).grid(row=0, column=0, padx=5)
        ttk.Button(ders_buton_frame, text="Ders Duzenle", 
                  command=self.ders_duzenle).grid(row=0, column=1, padx=5)

    def ders_kaydet(self):
        """Ders bilgilerini kaydet"""
        kod = self.ders_kodu.get().strip()
        ad = self.ders_adi.get().strip()
        
        if not kod or not ad:
            messagebox.showerror("Hata", "Ders kodu ve adı boş bırakılamaz!")
            return
        
        self.dersler.append((kod, ad))
        messagebox.showinfo("Başarılı", "Ders kaydedildi!")

    def ders_duzenle(self):
        # Düzenleme işlemi burada yapılabilir
        messagebox.showinfo("Bilgi", "Ders düzenleme işlemi burada yapılabilir.")

    def tablo_temizle(self, tablo_adi):
        """Seçilen tabloyu temizle"""
        if tablo_adi == "tablo1":
            self.tablo1_data = None
            self.durum_etiketi_guncelle(self.tablo1_label, False)
        elif tablo_adi == "tablo2":
            self.tablo2_data = None
            self.durum_etiketi_guncelle(self.tablo2_label, False)
        elif tablo_adi == "notlar":
            self.notlar_data = None
            self.durum_etiketi_guncelle(self.notlar_label, False)
        messagebox.showinfo("Bilgi", f"{tablo_adi.capitalize()} temizlendi.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DersBasariHesaplamaArayuz(root)
    root.mainloop()
