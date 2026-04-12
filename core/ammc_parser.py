# ═══════════════════════════════════════════════════════════════
# FIX _detect_val_cols — les 4 dernières lignes du CPC (XIV/XV/XVI)
# sont dans un tableau séparé de 3 lignes sur la page 5.
# L'ancienne version skippait les 2 premières lignes → ne restait
# qu'1 ligne → aucune colonne n'atteignait le seuil de 2 → val_cols=[]
#
# Fix : adapter le skip et le seuil à la taille du tableau.
# ═══════════════════════════════════════════════════════════════

# Dans ammc_parser.py, remplacer:

# ANCIEN CODE:
# def _detect_val_cols(table) -> list:
#     hits = {}
#     for row in (table or [])[2:]:
#         if not row: continue
#         for ci, c in enumerate(row):
#             if not c: continue
#             s = str(c).strip().replace('\xa0','').replace(' ','')
#             if re.match(r'^-?\d{1,3}(\.\d{3})*,\d{2}$', s) or re.match(r'^-?\d+,\d{2}$', s):
#                 hits[ci+1] = hits.get(ci+1, 0) + 1
#     return sorted(k for k, v in hits.items() if v >= 2)

# NOUVEAU CODE:
def _detect_val_cols(table) -> list:
    hits = {}
    tbl = table or []
    # Adapter le skip selon la taille du tableau :
    # - tableaux normaux (>= 5 lignes) : skip 2 lignes de header
    # - petits tableaux 3-4 lignes (ex: XIV/XV/XVI en fin de CPC) : skip 0
    skip = 2 if len(tbl) >= 5 else 0
    for row in tbl[skip:]:
        if not row: continue
        for ci, c in enumerate(row):
            if not c: continue
            s = str(c).strip().replace('\xa0','').replace(' ','')
            if re.match(r'^-?\d{1,3}(\.\d{3})*,\d{2}$', s) or re.match(r'^-?\d+,\d{2}$', s):
                hits[ci+1] = hits.get(ci+1, 0) + 1
    # Seuil adapté : >= 2 pour les grands tableaux, >= 1 pour les petits
    min_hits = 2 if len(tbl) >= 5 else 1
    return sorted(k for k, v in hits.items() if v >= min_hits)
