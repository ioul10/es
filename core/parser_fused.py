"""
core/parser_fused.py
Parser pour PDFs Style 3 : cellules multi-postes fusionnées (ex: SGTM).
Complète les valeurs X/Y en utilisant les lignes normales de extract_tables.
"""
import re, unicodedata
import pdfplumber

def _parse(s):
    """Parse un nombre au format français."""
    if not s: return None
    s = str(s).strip().replace('\xa0', '').replace(' ', '')
    s = s.replace('\'', '').replace('\u202f', '')
    s = re.sub(r'[^\d,.-]', '', s)
    if not s: return None
    s = s.replace(',', '.')
    if s.count('.') > 1:
        parts = s.rsplit('.', 1)
        s = parts[0].replace('.', '') + '.' + parts[1]
    try:
        v = float(s)
        return None if abs(v) > 1e13 else v
    except: return None


def _is_fused_table(table) -> bool:
    """Détecte si un tableau a des cellules fusionnées multi-valeurs."""
    fused_count = 0
    for row in table[2:]:
        if not row: continue
        for cell in row:
            if cell and '\n' in str(cell):
                parts = str(cell).split('\n')
                num_parts = sum(1 for p in parts if _parse(p.strip()) is not None)
                if num_parts >= 2:
                    fused_count += 1
    return fused_count >= 3


def _extract_normal_rows(table) -> list:
    """
    Extrait les lignes normales (non fusionnées) avec 4 valeurs numériques.
    Ce sont les totaux de section dans le PDF SGTM.
    """
    rows = []
    for row in table:
        if not row: continue
        # Vérifier que la colonne de valeurs n'est pas fusionnée
        val_col = str(row[2]) if len(row) > 2 and row[2] else ''
        if '\n' in val_col: continue
        # Extraire les 4 colonnes de valeurs
        vals = [_parse(str(row[i])) if i < len(row) and row[i] else None 
                for i in [2, 3, 4, 5]]
        if any(v is not None and v != 0 for v in vals):
            rows.append(vals)
    return rows


def _complete_with_table(xy_rows: list, et_rows: list, tol: float = 0.02) -> list:
    """
    Pour chaque ligne X/Y avec valeurs incomplètes (< 4),
    chercher dans et_rows une ligne contenant une valeur commune.
    Si trouvée → remplacer par les 4 valeurs complètes.
    """
    used = [False] * len(et_rows)
    result = []

    for label, vals in xy_rows:
        known = [v for v in vals if v is not None and v != 0]

        if len(known) >= 4:
            result.append((label, vals[:4]))
            continue

        if not known:
            result.append((label, [None, None, None, None]))
            continue

        # Chercher la meilleure correspondance dans et_rows
        best_match = None
        best_score = 0

        for i, et_vals in enumerate(et_rows):
            if used[i]: continue
            et_known = [v for v in et_vals if v is not None and v != 0]
            if not et_known: continue

            score = sum(
                1 for kv in known for ev in et_known
                if abs(kv - ev) < tol * max(abs(kv), abs(ev), 1)
            )

            if score > best_score:
                best_score = score
                best_match = i

        if best_match is not None and best_score > 0:
            used[best_match] = True
            result.append((label, et_rows[best_match]))
        else:
            v4 = (vals + [None, None, None, None])[:4]
            result.append((label, v4))

    return result


def extract_fused_section(pdf, page_indices: list) -> list:
    """
    Extrait une section d'un PDF Style 3 (cellules fusionnées SGTM).
    Combine X/Y (labels + valeurs partielles) avec extract_tables (valeurs complètes).
    Retourne list of (label, [brut_ou_exN, amort, netN, netN1]).
    """
    all_rows = []

    for idx in page_indices:
        if idx >= len(pdf.pages):
            continue
        page = pdf.pages[idx]

        # Import lazy anti-circular
        from core.ammc_parser import _xy_rows, _detect_val_cols
        # 1. X/Y → labels + valeurs partielles
        xy = _xy_rows(page)

        # 2. extract_tables → lignes normales avec valeurs complètes
        tables = page.extract_tables()
        good = [t for t in tables
                if t and len(t) >= 2 and t[0] and len(t[0]) >= 4
                and sum(1 for r in t[1:] if any(c for c in r if c)) >= 2]

        if not good:
            all_rows.extend([(l, (v + [None,None,None,None])[:4]) for l, v in xy])
            continue

        tab = good[0]
        et_rows = _extract_normal_rows(tab)

        # 3. Compléter X/Y avec les valeurs de extract_tables
        completed = _complete_with_table(xy, et_rows)
        all_rows.extend(completed)

    return all_rows
