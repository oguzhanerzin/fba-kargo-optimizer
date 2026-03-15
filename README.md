# FBA Kargo Optimizer

Amazon FBA sevkiyatları için OR-Tools CP-SAT solver ile optimal kargo planı.

## Özellikler
- OR-Tools CP-SAT ile matematiksel optimal çözüm
- Gerçek ağırlık hard limit (46 lb)
- SKU/koli dengesi garantili
- Hacimsel ağırlık dengesi
- Stok bazlı bölünme limiti (≤10→2, 11-20→3, ≥21→4)
- Duplicate ASIN otomatik birleştirilir
- 4 sayfalı Excel çıktı (Veri, Koli_Ozet, Koli_Detay, Box_Product_Summary)

## Render Deploy (5 dakika)

1. Bu repoyu GitHub'a push et
2. [render.com](https://render.com) → New → Web Service
3. GitHub reposunu bağla
4. **Build Command:** `pip install -r requirements.txt`
5. **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
6. Deploy → URL'yi al

## Excel Format

Veri sayfası şu sütunları içermeli:
| ASIN | Ürün Adı | Birim Ağırlık (lb) | Birim Hacimsel Ağ. (lbs) | Adet |

## n8n Entegrasyonu

```
HTTP Request node:
  Method: POST
  URL: https://your-app.onrender.com/optimize
  Body: Form-Data
    file: [Excel dosyası]
  Response: Binary (Excel dosyası)
```

## Local Çalıştırma

```bash
pip install -r requirements.txt
python app.py
# → http://localhost:5000
```
