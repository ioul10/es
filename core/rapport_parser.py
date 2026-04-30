"""
core/rapport_parser.py
Parser pour rapports financiers libres (non-DGI, non-AMMC).
L'utilisateur indique manuellement les pages contenant Actif, Passif, CPC.
Supporte :
  - PDF texte normal (tables pdfplumber standards)
  - PDF avec texte rotatif (type LabelVie) → reconstruction par mots x/y
"""
import re
from collections import defaultdict
import pdfplumber

from core.dgi_parser import (
    ACTIF, PASSIF, CPC,
    _parse, _has_fused, _detect_val_cols, _is_cpc_6col,
    _xy_rows, match_label, _build_value_map,
    _is_index_row, _is_rotated,
)
_ROMAIN_RE = re.compile(r'^(I{1,3}|IV|V?I{0,3}|IX|X{1,3}|XIV|XV|XVI)$', re.I)


# ══════════════════════════════════════════════════════════════════
# UTILITAIRES
# ══════════════════════════════════════════════════════════════════

def _parse_pages_input(raw: str) -> list[int]:
    """Convertit saisie utilisateur en indices 0-based. Ex: '1,2' → [0,1]"""
    indices = []
    for part in re.split(r'[,;]', raw.strip()):
        part = part.strip()
        if '-' in part:
            bounds = part.split('-')
            try:
                a, b = int(bounds[0].strip()), int(bounds[1].strip())
                indices.extend(range(a - 1, b))
            except: pass
        else:
            try:
                indices.append(int(part) - 1)
            except: pass
    return sorted(set(i for i in indices if i >= 0))


def _is_rotated_page(page) -> bool:
    """Détecte si la majorité des mots de la page sont rotatifs (ttb)."""
    words = page.extract_words(x_tolerance=5, y_tolerance=3)
    if not words: return False
    rotated = sum(1 for w in words if w.get('direction') != 'ltr')
    return rotated / len(words) > 0.5


def _reconstruct_tokens(word_list: list) -> list[str]:
    """
    Fusionne les fragments de mots espacés en tokens complets.
    - gap < FRAG_THRESH : fragment du même mot → coller sans espace
    - gap < WORD_THRESH : mots proches → nouveau token
    - gap >= WORD_THRESH : tokens séparés
    """
    if not word_list: return []
    FRAG_THRESH = 2   # px : en-dessous → coller (fragment de lettre)
    sorted_words = sorted(word_list, key=lambda w: w['x0'])
    tokens = []
    current = sorted_words[0]['text']
    current_x1 = sorted_words[0]['x1']

    for w in sorted_words[1:]:
        gap = w['x0'] - current_x1
        if gap < FRAG_THRESH:   # fragment de lettre → coller
            current += w['text']
        else:                   # nouveau token (avec espace si besoin)
            tokens.append(current)
            current = w['text']
        current_x1 = max(current_x1, w['x1'])
    tokens.append(current)
    return tokens


def _clean_label(label: str) -> str:
    """Nettoie un label extrait : retire *, espaces multiples, etc."""
    label = re.sub(r'^\*\s*', '', label)
    label = re.sub(r'\s+', ' ', label).strip()
    return label


# ══════════════════════════════════════════════════════════════════
# EXTRACTION PAGE NORMALE (tables pdfplumber)
# ══════════════════════════════════════════════════════════════════

def _extract_normal(page, is_actif: bool = False) -> list[tuple[str, list]]:
    """Extraction standard via pdfplumber.extract_tables()."""
    tables = page.extract_tables()
    good = [t for t in tables
            if t and len(t) >= 3 and t[0] and len(t[0]) >= 2
            and sum(1 for r in t[1:] if any(c for c in r if c)) >= 2]
    if not good: return []

    rows = []
    for t in good:
        if _is_cpc_6col(t):
            for row in t[1:]:
                if not row: continue
                col0  = str(row[0] or '').strip().replace('\n', ' ')
                col1  = str(row[1] or '').strip().replace('\n', ' ')
                label = _clean_label(col1 if col1 else col0)
                if not label or len(label) < 2: continue
                if _parse(label) is not None: continue
                v = [
                    _parse(row[2]) if len(row) > 2 else None,
                    _parse(row[3]) if len(row) > 3 else None,
                    _parse(row[4]) if len(row) > 4 else None,
                    _parse(row[5]) if len(row) > 5 else None,
                ]
                if any(x is not None for x in v):
                    rows.append((label, v))
            continue

        val_cols = _detect_val_cols(t)
        for row in t:
            if not row or _is_index_row(row): continue
            if len(row) < 2: continue
            label = ''
            for ci in [0, 1]:
                if ci < len(row) and row[ci]:
                    candidate = _clean_label(str(row[ci]).replace('\n', ' '))
                    if (candidate and len(candidate) >= 2
                            and _parse(candidate) is None
                            and not _is_rotated(candidate)
                            and not _ROMAIN_RE.match(candidate.strip())):
                        label = candidate
                        break
            if not label: continue

            def gv(col):
                if col and len(row) >= col: return _parse(row[col - 1])
                return None

            nv = len(val_cols)
            if is_actif and nv == 3:
                gap = val_cols[1] - val_cols[0]
                vals = ([gv(val_cols[0]), None, gv(val_cols[1]), gv(val_cols[2])]
                        if gap > 1 else
                        [gv(val_cols[0]), gv(val_cols[1]), gv(val_cols[2]), None])
            else:
                vals = [gv(val_cols[i]) for i in range(min(nv, 4))]

            if any(v is not None for v in vals):
                rows.append((label, vals))

    return rows


# ══════════════════════════════════════════════════════════════════
# EXTRACTION PAGE ROTATIVE (mots x/y)
# ══════════════════════════════════════════════════════════════════

def _extract_rotated_zone(page, x_min: float, x_max: float,
                           n_val_cols: int) -> list[tuple[str, list]]:
    """
    Extrait les lignes d'une zone x_min..x_max d'une page rotative.
    Utilise x_tolerance=8 pour que pdfplumber fusionne déjà les fragments
    de lettres → les mots arrivent bien formés sans reconstruction manuelle.
    """
    # x_tolerance=8 fusionne les fragments de lettres espacées (style LabelVie)
    words = page.extract_words(x_tolerance=8, y_tolerance=5)
    zone  = [w for w in words
             if w.get('direction') != 'ltr'
             and x_min <= w['x0'] < x_max]
    if not zone: return []

    # Grouper par ligne (top) avec tolérance 8px
    lines = defaultdict(list)
    for w in zone:
        key = round(w['top'] / 8) * 8
        lines[key].append(w)

    rows = []
    skip_patterns = [
        r'^(brut|amort\.?|net|exercice|désignation|bilan|actif$|passif$|nature$)$',
        r'^\d{2}/\d{2}/\d{4}',
        r'^(31/12|01/01)',
    ]

    for y in sorted(lines.keys()):
        wlist = sorted(lines[y], key=lambda w: w['x0'])

        label_parts = []
        raw_tokens  = []   # (x0, text) pour reconstituer les nombres fragmentés

        for w in wlist:
            tok = w['text'].strip()
            if not tok: continue
            if len(tok) > 1 and _parse(tok) is None and not _ROMAIN_RE.match(tok):
                label_parts.append(tok)
            else:
                raw_tokens.append((w['x0'], tok))

        label = _clean_label(' '.join(label_parts))
        if not label or len(label) < 2: continue
        if any(re.match(p, label.lower()) for p in skip_patterns): continue

        # Reconstituer les nombres fragmentés : "1" "639" "931" "012,84" → 1639931012.84
        val_tokens = []
        i = 0
        while i < len(raw_tokens):
            x0, tok = raw_tokens[i]
            # Accumuler les fragments numériques proches (gap < 25px)
            accumulated = tok
            j = i + 1
            while j < len(raw_tokens):
                nx0, ntok = raw_tokens[j]
                gap = nx0 - (x0 + len(accumulated) * 4)  # estimation largeur
                # Si le prochain token est purement numérique ET proche → fragment
                if (gap < 30 and
                    re.match(r'^\d', ntok) and
                    not re.search(r',\d{2}$', accumulated)):
                    accumulated += ntok
                    j += 1
                else:
                    break
            v = _parse(accumulated)
            if v is not None:
                val_tokens.append(v)
            i = j if j > i + 1 else i + 1

        if val_tokens:
            vals = val_tokens[:n_val_cols]
            vals += [None] * (n_val_cols - len(vals))
            rows.append((label, vals))

    return rows


def _find_zone_boundary(page) -> float:
    """
    Sur une page rotative avec Actif+Passif côte à côte,
    détecte la frontière X entre les deux zones.
    Cherche le 'gap' le plus large entre les clusters de x0.
    """
    words = page.extract_words(x_tolerance=8, y_tolerance=5)
    rotated = [w for w in words if w.get('direction') != 'ltr']
    if not rotated: return page.width / 2

    xs = sorted(set(round(w['x0'] / 10) * 10 for w in rotated))
    if len(xs) < 2: return page.width / 2

    # Chercher le plus grand gap entre x consécutifs
    max_gap = 0
    boundary = page.width / 2
    for i in range(1, len(xs)):
        gap = xs[i] - xs[i-1]
        if gap > max_gap and 200 < xs[i] < 600:
            max_gap = gap
            boundary = (xs[i] + xs[i-1]) / 2

    return boundary


# ══════════════════════════════════════════════════════════════════
# EXTRACTION SECTION (dispatching normal vs rotatif)
# ══════════════════════════════════════════════════════════════════

def _extract_section(pdf, page_indices: list,
                     section: str,           # 'actif', 'passif', 'cpc'
                     ) -> list[tuple[str, list]]:
    """
    Extrait les lignes d'une section depuis les pages indiquées.
    Détecte automatiquement si la page est rotative ou normale.
    """
    n_val = {'actif': 4, 'passif': 2, 'cpc': 4}[section]
    is_actif = section == 'actif'
    all_rows = []

    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        page = pdf.pages[idx]

        if _is_rotated_page(page):
            # Page rotative → extraction par zones x
            boundary = _find_zone_boundary(page)

            if section == 'actif':
                rows = _extract_rotated_zone(page, 0, boundary, n_val)
            elif section == 'passif':
                rows = _extract_rotated_zone(page, boundary, page.width, n_val)
            else:  # cpc — toute la largeur
                rows = _extract_rotated_zone(page, 0, page.width, n_val)
        else:
            # Page normale → extraction via tables
            if _has_fused(page.extract_tables()[0] if page.extract_tables() else []):
                rows = _xy_rows(page)
            else:
                rows = _extract_normal(page, is_actif=is_actif)

        all_rows.extend(rows)

    return all_rows


# ══════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════

def parse(pdf_path: str, pages_actif: str, pages_passif: str,
          pages_cpc: str, info: dict) -> dict:
    """
    Parse un rapport financier libre.
    Retourne un dict compatible avec excel_writer.write().
    """
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)

    idx_actif  = _parse_pages_input(pages_actif)
    idx_passif = _parse_pages_input(pages_passif)
    idx_cpc    = _parse_pages_input(pages_cpc)

    actif_rows  = _extract_section(pdf, idx_actif,  'actif')
    passif_rows = _extract_section(pdf, idx_passif, 'passif')
    cpc_rows    = _extract_section(pdf, idx_cpc,    'cpc')
    pdf.close()

    actif_map  = _build_value_map(actif_rows,  ACTIF)
    passif_map = _build_value_map(passif_rows, PASSIF)
    cpc_map    = _build_value_map(cpc_rows,    CPC)

    info.setdefault('format', 'Rapport')

    return {
        'info':      info,
        'actif':     actif_map,
        'passif':    passif_map,
        'cpc':       cpc_map,
        'format':    'Rapport',
        'templates': {'actif': ACTIF, 'passif': PASSIF, 'cpc': CPC},
        '_stats': {
            'total_pages': total_pages,
            'actif':       len(actif_map),
            'passif':      len(passif_map),
            'cpc':         len(cpc_map),
            'actif_max':   len(ACTIF),
            'passif_max':  len(PASSIF),
            'cpc_max':     len(CPC),
        }
    }
