# planlar

Ogretmenler icin yillik plan Excel dosyalari. TAP uygulamasi buradan yalniz indirir.

## Dosya duzeni

Klasor = ders. Dosya adi = okul turu.

```text
Fizik/Fizik Anadolu Liseleri.xlsx
Fizik/FEN LISESI FIZIK.xlsx
```

## Otomatik katalog

Her `.xlsx` yuklemesinde GitHub Action `catalog.json` uretir.

- Katalog: https://raw.githubusercontent.com/mutludurmaz/planlar/main/catalog.json
- Ilk kurulum: Settings → Actions → General → Workflow permissions → Read and write

## Senaryolar (ortak yazili konu-soru dagilimi)

MEB ODSGM'nin yayimladigi konu-soru dagilim tablolari degistirilmeden konur;
TAP > Ogretmen Dosyalari > Senaryolar buradan okur.

```text
senaryolar/2026-2027-1/Fizik.xlsx   (yil-yil-donem / Ders.xlsx)
```

- Klasor adi `yyyy-yyyy-1` veya `yyyy-yyyy-2`; en yeni donem uygulamada acilir.
- Dosya adi TAP'taki ders adiyla eslesir (Fizik, Kimya, Turk Dili ve Edebiyati...).
- Bu klasordeki dosyalar yillik plan katalogunda gorunmez; `catalog.json` icinde
  ayri `senaryolar` bolumune yazilir.
- Kaynak: https://odsgm.meb.gov.tr (donem basi "Konu Soru Dagilim Tablolari" duyurusu)
