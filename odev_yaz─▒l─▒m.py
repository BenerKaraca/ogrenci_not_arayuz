import pandas as pd
import numpy as np
# Tablo 1: Program Çıktıları ve Ders Çıktıları İlişki Matrisi
try:
    tablo1 = pd.read_excel('Tablo1.xlsx', sheet_name=0)  # İlk sayfayı oku
    print("Tablo 1 başarıyla okundu.")
except Exception as e:
    print(f"Tablo1.xlsx dosyası okunurken hata oluştu: {e}")
    exit()

# Tablo 2: Ders Çıktıları ve Değerlendirme Kriterleri İlişki Matrisi
try:
    tablo2 = pd.read_excel('Tablo2.xlsx', sheet_name=0)  # İlk sayfayı oku
    print("Tablo 2 başarıyla okundu.")
except Exception as e:
    print(f"Tablo2.xlsx dosyası okunurken hata oluştu: {e}")
    exit()

# NotYukle: Öğrenci Notları
try:
    notlar = pd.read_excel('NotYukle.xlsx', sheet_name=0)  # İlk sayfayı oku
    print("NotYukle.xlsx dosyası başarıyla okundu.")
except Exception as e:
    print(f"NotYukle.xlsx dosyası okunurken hata oluştu: {e}")
    exit()

# Tablo 1'den Program Çıktıları ve İlişki Değerlerini Al
try:
    # İlişki matrisini 0 ve 1'lere yuvarlayalım
    program_ciktilari_iliski = tablo1.iloc[0:10, 1:6].values  # 0'dan 9'a kadar (10 satır)
    program_ciktilari_iliski = np.round(program_ciktilari_iliski).astype(int)  # Ondalık değerleri yuvarla
    
    if program_ciktilari_iliski.shape[0] != 10:
        raise ValueError(f"Program çıktıları sayısı 10 olmalı, bulunan: {program_ciktilari_iliski.shape[0]}")
    
    iliski_degerleri = tablo1.iloc[0:10, -1].values
    print("Tablo 1 verileri başarıyla alındı.")
except Exception as e:
    print(f"Tablo 1 verileri alınırken hata oluştu: {e}")
    exit()

# Tablo 2'den Ders Çıktıları ve Değerlendirme Kriterlerini Al
try:
    # İlk satır başlıklar, ikinci satır yüzdeler, sonraki satırlar ilişki matrisi
    iliski_matrisi = tablo2.iloc[1:7, 1:6].values  # Tüm ders çıktıları (5 satır) dahil edildi
    oranlar = [10, 10, 10, 30, 40]  # Sabit yüzdeler: Ödev1, Ödev2, Quiz, Vize, Final
    print("Tablo 2 verileri başarıyla alındı.")
except Exception as e:
    print(f"Tablo 2 verileri alınırken hata oluştu: {e}")
    exit()

# NotYukle'den Öğrenci Notlarını ve Numaralarını Al
try:
    ogrenci_notlari = notlar.iloc[:, 1:7].values  # Öğrenci notları (Ödev1, Ödev2, Quiz, Vize, Final)
    ogrenci_nolar = notlar['Ogrenci_No'].values  # Öğrenci numaraları
    print("NotYukle verileri başarıyla alındı.")
except Exception as e:
    print(f"NotYukle verileri alınırken hata oluştu: {e}")
    exit()

# Ders Çıktılarının String Değerleri
ders_ciktilari = [f'Ders Çıktı {i + 1}' for i in range(len(iliski_matrisi))]

# Tablo 3: Ağırlıklı Değerlendirme Tablosu Oluştur
agirlikli_degerlendirme = []
toplamlar = []

for i in range(len(iliski_matrisi)):
    agirlikli_satir = []
    toplam = 0
    for j in range(len(iliski_matrisi[i])):
        # İlişki değeri (0 veya 1) ile değerlendirme kriterinin yüzdesini çarp
        iliski = float(iliski_matrisi[i][j])
        yuzde = float(oranlar[j]) / 100
        deger = iliski * yuzde
        agirlikli_satir.append(deger)
        toplam += deger
    agirlikli_degerlendirme.append(agirlikli_satir)
    toplamlar.append(toplam)

# Tablo 3'ü DataFrame'e Dönüştür
df_tablo3 = pd.DataFrame(agirlikli_degerlendirme, 
                        columns=['Ödev1', 'Ödev2', 'Quiz', 'Vize', 'Final'])
df_tablo3.insert(0, 'Ders Çıktı', [f'Ders Çıktı {i+1}' for i in range(len(iliski_matrisi))])
df_tablo3['Toplam'] = toplamlar

# Her Öğrenci İçin Ders Çıktıları Başarı Oranlarını Hesapla (Tablo 4)
ders_ciktilari_basari_oranlari = []
for ogrenci in ogrenci_notlari:
    basari_oranlari = []
    for i in range(len(iliski_matrisi)):
        toplam = 0
        for j in range(len(iliski_matrisi[i])):
            # İlişki matrisi (0/1) * öğrenci notu * değerlendirme kriteri yüzdesi
            deger = iliski_matrisi[i][j] * (ogrenci[j]/10) * (oranlar[j] / 100)
            toplam += deger

        # Max değeri hesaplama: İlişki matrisi satırındaki 1'lerin toplam ağırlığı * 100
        max_not = sum(iliski_matrisi[i][j] * (oranlar[j] / 100) for j in range(len(iliski_matrisi[i]))) * 10

        # Başarı yüzdesi hesaplama
        yuzde_basari = (toplam / max_not) * 100 if max_not != 0 else 0

        basari_oranlari.append({
            'Ders Çıktı': ders_ciktilari[i],
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

# Her Öğrenci İçin Program Çıktıları Başarı Oranlarını Hesapla (Tablo 5)
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
            ders_basarilari.append(basari_oranlari[j]['% Başarı'] if iliski_degeri > 0 else 0.0)
        
        basari_orani = round(toplam_basari / toplam_iliski, 1) if toplam_iliski > 0 else 0.0

        program_basari_oranlari.append({
            'Program Çıktı': i + 1,
            'Ders Çıktıları': ders_basarilari,
            'Başarı Oranı': basari_orani
        })
    program_ciktilari_basari_oranlari.append(program_basari_oranlari)

# Sonuçları Excel'e Yaz
try:
    with pd.ExcelWriter('Sonuclar.xlsx') as writer:
        # Tablo 3: Ağırlıklı Değerlendirme Tablosu
        df_tablo3.to_excel(writer, sheet_name='Tablo3', index=False)

        # Her Öğrenci İçin Tablo 4 ve Tablo 5
        for idx, ogrenci_no in enumerate(ogrenci_nolar):
            # Tablo 4: Ders Çıktıları Başarı Oranları
            df_tablo4 = pd.DataFrame(ders_ciktilari_basari_oranlari[idx])
            df_tablo4.to_excel(writer, sheet_name=f'Öğrenci {ogrenci_no} Tablo4', index=False)

            # Tablo 5: Program Çıktıları Başarı Oranları
            basari_data = []
            for prog in program_ciktilari_basari_oranlari[idx]:
                row = [prog['Program Çıktı']] + prog['Ders Çıktıları'] + [prog['Başarı Oranı']]
                basari_data.append(row)
            
            # Ders çıktılarını numaralandır
            columns = ['Prg Çıktı'] + [f'Ders çıktısı {i}' for i in range(1, 6)] + ['Başarı Oranı']
            
            df_tablo5 = pd.DataFrame(basari_data, columns=columns)
            df_tablo5.to_excel(writer, sheet_name=f'Öğrenci {ogrenci_no} Tablo5', index=False)

    print("Hesaplamalar tamamlandı ve 'Sonuclar.xlsx' dosyasına kaydedildi.")
except Exception as e:
    print(f"Excel dosyasına yazılırken hata oluştu: {e}")
