"""
Amazon FBA Kargo Optimizasyon Motoru
OR-Tools CP-SAT solver ile optimal bin packing
"""
from ortools.sat.python import cp_model
import math


MAKS_GERCEK    = 46.0
MAKS_HACIMSEL  = 64.0
IDEAL_HACIMSEL = 49.0
SCALE          = 1000   # float → integer (OR-Tools integer ister)


def optimize(sku_listesi: list[dict]) -> dict:
    """
    sku_listesi: [{'asin', 'ad', 'gercek', 'hacimsel', 'adet'}, ...]
    Döner:  {'koliler': [...], 'stats': {...}}
    """
    # ── Duplicate ASIN birleştir ─────────────────────────────────────────
    asin_map = {}
    dup_list = []
    for s in sku_listesi:
        asin = s['asin'].strip()
        if not asin or not s['adet']: continue
        if asin in asin_map:
            asin_map[asin]['adet'] += int(s['adet'])
            if asin not in dup_list: dup_list.append(asin)
        else:
            asin_map[asin] = {
                'asin'     : asin,
                'ad'       : str(s.get('ad', '')).strip(),
                'gercek'   : float(str(s['gercek']).replace(',', '.')),
                'hacimsel' : float(str(s['hacimsel']).replace(',', '.')),
                'adet'     : int(s['adet']),
            }
    skular = list(asin_map.values())

    if not skular:
        return {'error': 'Geçerli ürün bulunamadı'}

    toplam_gercek   = sum(s['gercek']   * s['adet'] for s in skular)
    toplam_hacimsel = sum(s['hacimsel'] * s['adet'] for s in skular)
    toplam_adet     = sum(s['adet'] for s in skular)
    toplam_sku      = len(skular)
    min_koli        = math.ceil(toplam_gercek / MAKS_GERCEK)

    # SKU bölünme limiti
    def maks_bol(adet):
        if adet <= 10: return 2
        if adet <= 20: return 3
        return 4

    # ── CP-SAT modeli ─────────────────────────────────────────────────────
    model = cp_model.CpModel()
    N = toplam_sku        # SKU sayısı
    K = min_koli          # koli sayısı (sabit, minimize edilmiş)

    # x[i][k] = SKU i'den koli k'ya kaç adet gider
    x = [[model.NewIntVar(0, skular[i]['adet'], f'x_{i}_{k}')
          for k in range(K)] for i in range(N)]

    # y[i][k] = SKU i koli k'da var mı? (binary)
    y = [[model.NewBoolVar(f'y_{i}_{k}')
          for k in range(K)] for i in range(N)]

    # ── Kısıtlar ─────────────────────────────────────────────────────────

    # 1. Her SKU'nun tüm adetleri dağıtılmalı
    for i in range(N):
        model.Add(sum(x[i][k] for k in range(K)) == skular[i]['adet'])

    # 2. x[i][k] > 0 ⟺ y[i][k] = 1
    for i in range(N):
        for k in range(K):
            model.Add(x[i][k] <= skular[i]['adet'] * y[i][k])
            model.Add(x[i][k] >= y[i][k])

    # 3. SKU bölünme limiti: bir SKU en fazla maks_bol(adet) koliye bölünebilir
    for i in range(N):
        model.Add(sum(y[i][k] for k in range(K)) <= maks_bol(skular[i]['adet']))

    # 4. Gerçek ağırlık limiti (integer ölçekli)
    for k in range(K):
        model.Add(
            sum(int(skular[i]['gercek'] * SCALE) * x[i][k] for i in range(N))
            <= int(MAKS_GERCEK * SCALE)
        )

    # 5. Hacimsel ağırlık limiti
    for k in range(K):
        model.Add(
            sum(int(skular[i]['hacimsel'] * SCALE) * x[i][k] for i in range(N))
            <= int(MAKS_HACIMSEL * SCALE)
        )

    # ── Hedef fonksiyonu ─────────────────────────────────────────────────
    # Minimize et: (SKU/koli farkı) * 100 + (hacimsel sapma)
    # SKU sayısı dengesi
    hedef_sku = int(toplam_sku / K)

    sku_sayilari = [sum(y[i][k] for i in range(N)) for k in range(K)]

    # Her koli için SKU sapmasını hesapla (abs yerine iki taraflı)
    sku_sap_pos = [model.NewIntVar(0, N, f'sp_{k}') for k in range(K)]
    sku_sap_neg = [model.NewIntVar(0, N, f'sn_{k}') for k in range(K)]
    for k in range(K):
        model.Add(sku_sayilari[k] - hedef_sku == sku_sap_pos[k] - sku_sap_neg[k])

    # Hacimsel denge: koliler arası hacimsel fark
    h_vals = [
        sum(int(skular[i]['hacimsel'] * SCALE) * x[i][k] for i in range(N))
        for k in range(K)
    ]
    h_max = model.NewIntVar(0, int(MAKS_HACIMSEL * SCALE), 'h_max')
    h_min = model.NewIntVar(0, int(MAKS_HACIMSEL * SCALE), 'h_min')
    for k in range(K):
        model.Add(h_max >= h_vals[k])
        model.Add(h_min <= h_vals[k])

    # Gerçek ağırlık dengesi
    g_vals = [
        sum(int(skular[i]['gercek'] * SCALE) * x[i][k] for i in range(N))
        for k in range(K)
    ]
    g_max = model.NewIntVar(0, int(MAKS_GERCEK * SCALE), 'g_max')
    g_min = model.NewIntVar(0, int(MAKS_GERCEK * SCALE), 'g_min')
    for k in range(K):
        model.Add(g_max >= g_vals[k])
        model.Add(g_min <= g_vals[k])

    # Adet dengesi
    a_vals = [sum(x[i][k] for i in range(N)) for k in range(K)]
    a_max = model.NewIntVar(0, toplam_adet, 'a_max')
    a_min = model.NewIntVar(0, toplam_adet, 'a_min')
    for k in range(K):
        model.Add(a_max >= a_vals[k])
        model.Add(a_min <= a_vals[k])

    # Bileşik hedef:
    # SKU dengesi en önemli (×300), ardından gerçek ağırlık doluluk (×2), hacimsel denge (×1)
    model.Minimize(
        300 * (sum(sku_sap_pos) + sum(sku_sap_neg)) +
        2   * (g_max - g_min) +
        1   * (h_max - h_min) +
        1   * (a_max - a_min)
    )

    # ── Solver ───────────────────────────────────────────────────────────
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    solver.parameters.num_search_workers  = 4
    solver.parameters.log_search_progress = False

    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {'error': f'Çözüm bulunamadı (status={status}). Kısıtları gevşetin.'}

    # ── Sonucu derle ─────────────────────────────────────────────────────
    koliler = []
    for k in range(K):
        koli = {
            'no'            : k + 1,
            'parcalar'      : [],
            'toplam_gercek' : 0.0,
            'toplam_hacimsel': 0.0,
            'sku_sayisi'    : 0,
            'adet_toplam'   : 0,
        }
        for i in range(N):
            adet = solver.Value(x[i][k])
            if adet > 0:
                sku = skular[i]
                koli['parcalar'].append({
                    'asin'    : sku['asin'],
                    'ad'      : sku['ad'],
                    'gercek'  : sku['gercek'],
                    'hacimsel': sku['hacimsel'],
                    'adet'    : adet,
                })
                koli['toplam_gercek']    += sku['gercek']    * adet
                koli['toplam_hacimsel']  += sku['hacimsel']  * adet
                koli['adet_toplam']      += adet
        koli['sku_sayisi'] = len(koli['parcalar'])
        koliler.append(koli)

    sku_sayilari_sonuc = [k['sku_sayisi'] for k in koliler]
    gercek_sonuc       = [k['toplam_gercek'] for k in koliler]
    hacimsel_sonuc     = [k['toplam_hacimsel'] for k in koliler]
    adet_sonuc         = [k['adet_toplam'] for k in koliler]

    stats = {
        'toplam_koli'       : K,
        'toplam_adet'       : toplam_adet,
        'toplam_sku'        : toplam_sku,
        'sku_min'           : min(sku_sayilari_sonuc),
        'sku_max'           : max(sku_sayilari_sonuc),
        'sku_fark'          : max(sku_sayilari_sonuc) - min(sku_sayilari_sonuc),
        'gercek_min'        : round(min(gercek_sonuc), 2),
        'gercek_max'        : round(max(gercek_sonuc), 2),
        'hacimsel_min'      : round(min(hacimsel_sonuc), 2),
        'hacimsel_max'      : round(max(hacimsel_sonuc), 2),
        'adet_fark'         : max(adet_sonuc) - min(adet_sonuc),
        'duplicate_asinler' : dup_list,
        'solver_status'     : 'optimal' if status == cp_model.OPTIMAL else 'feasible',
    }

    return {'koliler': koliler, 'stats': stats, 'skular': skular}
