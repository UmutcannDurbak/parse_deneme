# ÇOK DOSYA İŞLEME ÖZELLİĞİ - KULLANIM KILAVUZU

## 📋 Özellik Özeti
Artık aynı anda birden fazla CSV dosyasını seçip/sürükleyip tek seferde işleyebilirsiniz!

## 🎯 Nasıl Kullanılır?

### Yöntem 1: Dosya Seçici ile (Ctrl+Click)
1. "CSV Dosyası Seç" butonuna tıklayın
2. İlk dosyayı seçin
3. **Ctrl tuşuna basılı tutarak** diğer dosyaları tıklayın
4. "Aç" butonuna tıklayın
5. Tüm dosyalar sırayla otomatik işlenecektir

### Yöntem 2: Sürükle-Bırak (Drag & Drop)
1. Klasörden birden fazla CSV dosyasını seçin
2. Hepsini birlikte tutun
3. Uygulama penceresine sürükleyin
4. Bırakın - otomatik işleme başlar!

## 📊 İşlem Süreci

### Her Dosya İçin:
```
[1/4] İşleniyor: adana27.csv
────────────────────────────────
[INFO] İşlem başladı [ADANA]
[INFO] CSV Dosyası: adana27.csv
[STEP] Tatlı eşleştirme başlıyor [ADANA]...
[STEP] Donuk eşleştirme başlıyor [ADANA]...
[STEP] Lojistik eşleştirme başlıyor [ADANA]...
[1/4] ✅ Başarılı: adana27.csv
```

### Son Özet:
```
================================================================================
[ÖZET] Toplu İşlem Tamamlandı
================================================================================
✅ Başarılı: 4/4 dosya
```

## 💬 Pop-up Mesajı (Tek Sefer)

Tüm dosyalar işlendikten sonra **tek bir pop-up** açılır:

```
┌──────────────────────────────────┐
│ Toplu İşlem Başarılı             │
├──────────────────────────────────┤
│ Toplu işlem tamamlandı!          │
│                                  │
│ 📊 Özet:                         │
│ • Toplam: 4 dosya                │
│ • Başarılı: 4 dosya              │
│                                  │
│ (Hata varsa burada listelenir)   │
└──────────────────────────────────┘
```

## ✅ Avantajlar

1. **Zaman Kazancı**: 10 dosyayı 10 kez değil, 1 kez seçip işleyin
2. **Güvenli İşlem**: Her dosya sırayla işlenir, birbiri üzerine yazmaz
3. **Detaylı Log**: Her dosya için ayrı log kaydı
4. **Hata Yönetimi**: Bir dosya hata verse bile diğerleri işlenir
5. **Tek Özet**: Sonunda tek pop-up ile detaylı rapor

## ⚠️ Önemli Notlar

- **Dosyalar Sırayla İşlenir**: Birbirlerinin üzerine yazmaz
- **Her Dosya Ayrı Excel'e Yazar**: Veriler temiz şekilde aktarılır
- **Hata Durumu**: Bir dosyada hata olursa, diğerleri devam eder
- **Log Takibi**: Her dosyanın işlemi log penceresinde görülür

## 🔧 Teknik Detaylar

### Tek Dosya İşleme (Önceki Gibi):
- 1 dosya seçilirse → Normal işlem + pop-up

### Çoklu Dosya İşleme (Yeni):
- 2+ dosya seçilirse → Toplu işlem + tek özet pop-up
- Her dosya için ayrı log satırları
- Son durumda özet rapor

## 📝 Örnek Senaryo

**Gün Sonu İşlemi:**
1. 27.11 klasöründeki tüm CSV'leri seçin (Ctrl+A)
2. Hepsini birlikte uyglamaya sürükleyin
3. Kahvenizi içerken otomatik işlenmesini bekleyin ☕
4. Tek pop-up'ta tüm sonuçları görün
5. Bitti! Tüm şubeler işlenmiş ✅

## 🎉 Sonuç

Artık **günlük 20-30 CSV dosyasını** tek seferde işleyebilirsiniz!
Her dosya için tek tek "Aç" > "Bekle" > "Tamam" döngüsüne gerek yok.

**Toplu işlem = Verimlilik + Zaman Kazancı!**
