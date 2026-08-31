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
