"""
core/table_detector.py
Détection automatique des tableaux financiers dans un PDF.
Fonctionne pour pages normales ET pages rotatives (LabelVie, etc.)

Approche :
  1. Détecter les frontières X via lignes verticales longues
  2. Classifier chaque bande : LABELS (texte) ou VALEURS (numériques)
  3. Regrouper : tableau = bande(s) labels + bandes valeurs adjacentes
  4. Identifier le type (Actif/Passif/CPC) par mots-clés + nb colonnes
"""
import re, unicodedata
from collections import defaultdict, Counter
import pdfplumber


# ══════════════════════════════════════════════════════════════════
# MOTS-CLÉS D'IDENTIFICATION
# ══════════════════════════════════════════════════════════════════

_KEYWORDS = {
    'actif': [
        'immobilisations en non valeurs', 'actif immobilise',
        'total i', 'stocks', 'tresorerie actif', 'frais preliminaires',
        'creances actif', 'total general', 'amort', 'brut',
        'immobilisations corporelles', 'immobilisations financieres',
    ],
    'passif': [
        'capitaux propres', 'dettes de financement', 'report a nouveau',
        'emprunts obligataires', 'resultat net exercice', 'total des capitaux',
        'provisions durables', 'tresorerie passif', 'reserve legale',
        'capital social', 'autres reserves',
    ],
    'cpc': [
        'produits exploitation', 'charges exploitation', 'chiffre affaires',
        'resultat exploitation', 'ventes marchandises', 'dotations exploitation',
        'resultat courant', 'resultat financier', 'produits financiers',
        'charges financieres', 'resultat non courant', 'impots resultats',
    ],
    'esg': [
        'valeur ajoutee', 'excedent brut', 'marge brute',
        'capacite autofinancement', 'autofinancement', 'soldes gestion',
        'production exercice', 'consommation exercice',
    ],
}


def _norm(s: str) -> str:
    s = unicodedata.normalize('NFD', str(s)).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', s)).strip()


def _is_num_token(t: str) -> bool:
    t2 = t.replace(',','').replace('.','').replace('-','').replace(' ','')
    return t2.isdigit() and len(t2) >= 1


def _merge_close(values: list, threshold: float) -> list:
    merged = []
    for v in sorted(values):
        if merged and v - merged[-1] < threshold:
            merged[-1] = (merged[-1] + v) / 2
        else:
            merged.append(float(v))
    return merged


def _is_rotated_page(page) -> bool:
    words = page.extract_words(x_tolerance=5, y_tolerance=3)
    if not words: return False
    rot = sum(1 for w in words if w.get('direction') != 'ltr')
    return rot / len(words) > 0.5


# ══════════════════════════════════════════════════════════════════
# DÉTECTION DES FRONTIÈRES X
# ══════════════════════════════════════════════════════════════════

def _get_x_boundaries(page) -> list[float]:
    """Détecte les frontières X via lignes verticales longues."""
    W, H = page.width, page.height
    v_lines = [e for e in page.edges
               if abs(e['x0'] - e['x1']) < 2
               and (e['bottom'] - e['top']) > H * 0.15]
    if not v_lines:
        return [0.0, W]
    xs = [round(e['x0'] / 3) * 3 for e in v_lines]
    merged = _merge_close(xs, threshold=12)
    # Ajouter bords
    result = [0.0] + [x for x in merged if 5 < x < W - 5] + [W]
    return sorted(set(result))


# ══════════════════════════════════════════════════════════════════
# ANALYSE DES BANDES
# ══════════════════════════════════════════════════════════════════

def _analyze_bands(page, x_boundaries: list) -> list[dict]:
    """
    Analyse chaque bande X et la classifie :
    - 'label'  : beaucoup de texte, peu de nombres
    - 'values' : majorité de nombres, peu de texte
    - 'empty'  : peu de contenu
    """
    W, H = page.width, page.height
    is_rot = _is_rotated_page(page)
    tol_x = 8 if is_rot else 5

    words = page.extract_words(x_tolerance=tol_x, y_tolerance=5)
    if is_rot:
        words = [w for w in words if w.get('direction') != 'ltr']

    bands = []
    for i in range(len(x_boundaries) - 1):
        x0, x1 = x_boundaries[i], x_boundaries[i + 1]
        w_band = x1 - x0
        if w_band < W * 0.04:
            continue

        # Mots dans la bande
        if is_rot:
            bw = [w for w in words if x0 <= w['x0'] < x1]
        else:
            bw = [w for w in words if x0 <= w['x0'] < x1]

        if not bw:
            bands.append({'x0': x0, 'x1': x1, 'kind': 'empty',
                          'words': [], 'n_num_cols': 0})
            continue

        n_total = len(bw)
        n_num   = sum(1 for w in bw if _is_num_token(w['text']))
        ratio   = n_num / n_total if n_total > 0 else 0

        # Compter les colonnes numériques distinctes
        if n_num > 0:
            if is_rot:
                num_positions = [round(w['top'] / 12) * 12
                                 for w in bw if _is_num_token(w['text'])]
            else:
                num_positions = [round(w['x0'] / 12) * 12
                                 for w in bw if _is_num_token(w['text'])]
            col_counts  = Counter(num_positions)
            n_num_cols  = len([c for c in col_counts.values() if c >= 2])
        else:
            n_num_cols = 0

        kind = ('values' if ratio > 0.45 and n_num_cols >= 1
                else 'label' if n_total >= 4
                else 'empty')

        bands.append({
            'x0': x0, 'x1': x1,
            'kind': kind,
            'words': bw,
            'n_num_cols': n_num_cols,
            'n_total': n_total,
            'ratio_num': ratio,
        })

    return bands


# ══════════════════════════════════════════════════════════════════
# REGROUPEMENT EN TABLEAUX
# ══════════════════════════════════════════════════════════════════

def _group_into_tables(bands: list, page_width: float,
                        page_height: float) -> list[dict]:
    """
    Regroupe les bandes en tableaux :
    un tableau = 1+ bandes 'label' + 1+ bandes 'values' adjacentes.
    Une bande 'values' large (> 30% page) avec ratio_num < 0.80
    est traitée comme un tableau autonome (labels + valeurs intégrés).
    """
    tables = []
    i = 0
    while i < len(bands):
        b = bands[i]

        # Cas spécial : grande bande mixte (labels + valeurs ensemble)
        is_wide  = (b['x1'] - b['x0']) > page_width * 0.25
        is_mixed = (b['kind'] == 'values' and
                    b.get('ratio_num', 1.0) < 0.80 and
                    is_wide and b.get('n_total', 0) >= 15)
        if is_mixed:
            typ, score = _identify(b['words'], b['n_num_cols'])
            tables.append({
                'x0': b['x0'], 'x1': b['x1'],
                'words': b['words'],
                'n_val_cols': b['n_num_cols'],
                'label_bands': [], 'value_bands': [b],
            })
            i += 1
            continue

        if b['kind'] == 'label' and b['n_total'] >= 8:
            # Chercher les bandes valeurs adjacentes
            label_bands  = [b]
            values_bands = []
            j = i + 1
            while j < len(bands):
                nb = bands[j]
                if nb['kind'] == 'values':
                    values_bands.append(nb)
                elif nb['kind'] == 'label' and not values_bands:
                    label_bands.append(nb)  # labels multiples
                else:
                    break
                j += 1

            if values_bands:
                all_bands = label_bands + values_bands
                all_words = []
                for band in all_bands:
                    all_words.extend(band['words'])

                n_val_cols = max(vb['n_num_cols'] for vb in values_bands)

                tables.append({
                    'x0':        min(b2['x0'] for b2 in all_bands),
                    'x1':        max(b2['x1'] for b2 in all_bands),
                    'words':     all_words,
                    'n_val_cols': n_val_cols,
                    'label_bands': label_bands,
                    'value_bands': values_bands,
                })
                i = j
                continue
        i += 1

    return tables


# ══════════════════════════════════════════════════════════════════
# IDENTIFICATION DU TYPE
# ══════════════════════════════════════════════════════════════════

def _identify(words: list, n_val_cols: int) -> tuple[str, int]:
    """Identifie le type par mots-clés + heuristique colonnes."""
    text = _norm(' '.join(w['text'] for w in words))
    scores = {}
    for typ, kws in _KEYWORDS.items():
        score = sum(1 for kw in kws if _norm(kw) in text)
        if score > 0:
            scores[typ] = score

    # Heuristique colonnes :
    # Actif → 4 colonnes (Brut, Amort, Net N, Net N-1)
    # Passif → 2 colonnes (N, N-1)
    # CPC → 4 colonnes (Propres, Préc, Total N, Total N-1)
    if not scores:
        if n_val_cols >= 3:
            return 'actif', 0
        elif n_val_cols == 2:
            return 'passif', 0
        return 'inconnu', 0

    # Si scores proches entre actif et cpc → arbitrer par mots-clés spécifiques
    if 'actif' in scores and 'cpc' in scores:
        actif_specific = ['brut', 'amort', 'immobilisations corporelles', 'stocks']
        cpc_specific   = ['chiffre affaires', 'resultat courant', 'charges exploitation']
        a_bonus = sum(1 for kw in actif_specific if _norm(kw) in text)
        c_bonus = sum(1 for kw in cpc_specific   if _norm(kw) in text)
        if c_bonus > a_bonus:
            scores['cpc'] += 2
        elif a_bonus > c_bonus:
            scores['actif'] += 2

    best = max(scores, key=scores.get)
    return best, scores[best]


# ══════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════

def detect_tables(pdf_path: str) -> list[dict]:
    """
    Détecte et identifie les tableaux financiers dans toutes les pages.

    Retourne une liste de dicts :
    {
        'type'    : 'actif'|'passif'|'cpc'|'esg'|'inconnu',
        'score'   : int,
        'page'    : int (1-based),
        'pct_x0'  : int (0-100),
        'pct_x1'  : int (0-100),
        'n_val_cols': int,
        'n_words' : int,
    }
    """
    pdf    = pdfplumber.open(pdf_path)
    result = []

    for page_idx, page in enumerate(pdf.pages):
        W, H = page.width, page.height

        x_bounds = _get_x_boundaries(page)
        bands    = _analyze_bands(page, x_bounds)
        tables   = _group_into_tables(bands, W, H)

        for t in tables:
            typ, score = _identify(t['words'], t['n_val_cols'])
            result.append({
                'type':       typ,
                'score':      score,
                'page':       page_idx + 1,
                'x0':         t['x0'],
                'x1':         t['x1'],
                'pct_x0':     round(t['x0'] / W * 100),
                'pct_x1':     round(t['x1'] / W * 100),
                'n_val_cols': t['n_val_cols'],
                'n_words':    len(t['words']),
            })

    pdf.close()
    result.sort(key=lambda t: (t['page'], t['x0']))
    return result


def summarize(tables: list[dict]) -> dict:
    """Retourne le meilleur candidat pour chaque section."""
    summary = {}
    for typ in ['actif', 'passif', 'cpc', 'esg']:
        candidates = [t for t in tables if t['type'] == typ]
        if candidates:
            best = max(candidates, key=lambda t: (t['score'], t['n_words']))
            summary[typ] = best
    return summary
