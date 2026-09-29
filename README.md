# Excel-difference

İki Excel dosyasını (`.xlsx` / `.xlsm`) **tüm sheet'leriyle** karşılaştırıp farkları raporlayan küçük bir Python programı.

## Neleri karşılaştırır?

| Kategori | Detay |
|---|---|
| Sheet | Eklenen / silinen sheet, sheet sırası, görünürlük, sekme rengi, freeze panes, otomatik filtre, koşullu biçimlendirme / veri doğrulama sayısı, sheet koruması |
| Değer / Formül | Hücre içerikleri, formüller, formüllerin hesaplanmış (cache) değerleri, veri tipi |
| Font | Ad, boyut, kalın, italik, altı/üstü çizili, renk, üst/alt simge, anahat, gölge |
| Dolgu | Dolgu tipi, ön plan / arka plan rengi, gradient |
| Kenarlık | Sol / sağ / üst / alt / çapraz stil ve renkleri |
| Hizalama | Yatay, dikey, metni kaydır, sığdır, girinti, metin dönüşü |
| Sayı formatı | Örn. `0.00`, `dd.mm.yyyy`, `%` |
| Koruma | Kilitli / gizli |
| Diğer | Yorumlar, hyperlink'ler, birleştirilmiş hücreler, sütun genişlikleri, satır yükseklikleri, gizli satır/sütunlar, tanımlı isimler |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
# Sadece ekrana yaz
python excel_diff.py eski.xlsx yeni.xlsx

# Excel raporu üret (Ozet + Farklar sheet'leri)
python excel_diff.py eski.xlsx yeni.xlsx -o fark_raporu.xlsx

# CSV (; ayraçlı, Excel'de Türkçe karakterlerle açılır) veya TXT rapor
python excel_diff.py eski.xlsx yeni.xlsx -o fark_raporu.csv
python excel_diff.py eski.xlsx yeni.xlsx -o fark_raporu.txt
```

Seçenekler:

- `--no-format` : Format/font farklarını atla, sadece değerleri karşılaştır
- `--no-cached` : Formüllerin hesaplanmış değerlerini karşılaştırma
- `-q` / `--quiet` : Ekrana sadece toplam fark sayısını yaz

Çıkış kodu: `0` = fark yok, `1` = fark var, `2` = hata.

## Örnek çıktı

```
[Sheet: Veri]
  D1:E1      Birlestirme / Birlestirilmis hucre: yok  ->  var
  Sutun A    Sutun / Genislik: 13.0  ->  30.0
  A1         Dolgu / Tip: (bos)  ->  solid
  B1         Font / Ad: Calibri  ->  Arial
  B1         Font / Kalin: False  ->  True
  A2         Deger / Icerik: 10  ->  11
  C5         Yorum / Metin: (bos)  ->  not
```

## Notlar

- Eski `.xls` formatı desteklenmez; önce Excel'de `.xlsx` olarak kaydedin.
- Grafikler, resimler ve makro (VBA) kodu karşılaştırılmaz.
- Hesaplanmış değerler Excel'in dosyaya kaydettiği cache'ten okunur; dosya Excel dışında üretildiyse bu değerler boş olabilir.

## Test

```bash
python test_excel_diff.py
```
