"""
core/ammc_parser.py
Parser AMMC (5 pages) : extrait les valeurs et les mappe sur le template fixe MCN.
Template fixe = même structure pour tous les PDFs AMMC, quelle que soit la mise en page.
"""
import re, unicodedata
from collections import defaultdict
import pdfplumber
from utils.logger import get_logger
from core.synonyms import lookup_in_template

logger = get_logger(__name__)

_ROMAIN_RE = re.compile(r'^(XIV|XIII|XII|XI|IX|VIII|VII|VI|IV|XVI|XV|X|V|III|II|I)$', re.I)

# ══════════════════════════════════════════════════════════════════
# TEMPLATE FIXE MCN — AMMC
# (key_normalisée, label_affiché, type)
# ══════════════════════════════════════════════════════════════════

ACTIF = [
    ('immobilisations non valeurs',          'Immobilisations en non valeurs [A]',             'section'),
    ('frais preliminaires',                  'Frais préliminaires',                             'normal'),
    ('charges repartir',                     'Charges à répartir sur plusieurs exercices',      'normal'),
    ('primes remboursement obligations',     'Primes de remboursement des obligations',         'normal'),
    ('immobilisations incorporelles',        'Immobilisations incorporelles [B]',               'section'),
    ('immobilisations recherche',            'Immobilisations en Recherche et Développement',   'normal'),
    ('brevets marques droits',               'Brevets, marques, droits et valeurs similaires',  'normal'),
    ('fonds commercial',                     'Fonds commercial',                                'normal'),
    ('autres immobilisations incorporelles', 'Autres immobilisations incorporelles',            'normal'),
    ('immobilisations corporelles',          'Immobilisations corporelles [C]',                 'section'),
    ('terrains',                             'Terrains',                                        'normal'),
    ('constructions',                        'Constructions',                                   'normal'),
    ('installations techniques',             'Installations techniques, matériel et outillage', 'normal'),
    ('materiel transport',                   'Matériel de transport',                           'normal'),
    ('mobilier materiel bureau',             'Mobilier, Mat. de bureau, Aménagement divers',    'normal'),
    ('autres immobilisations corporelles',   'Autres immobilisations corporelles',              'normal'),
    ('immobilisations corporelles cours',    'Immobilisations corporelles en cours',            'normal'),
    ('immobilisations financieres',          'Immobilisations financières [D]',                 'section'),
    ('prets immobilises',                    'Prêts immobilisés',                               'normal'),
    ('autres creances financieres',          'Autres créances financières',                     'normal'),
    ('titres participation',                 'Titres de participation',                         'normal'),
    ('autres titres immobilises',            'Autres titres immobilisés',                       'normal'),
    ('ecarts conversion actif immobilise',   'Écarts de conversion actif [E]',                  'section'),
    ('diminution creances immobilisees',     'Diminution des créances immobilisées',            'normal'),
    ('augmentations dettes financement',     'Augmentations des dettes de financement',         'normal'),
    ('total i actif',                        'TOTAL I (A+B+C+D+E)',                             'total'),
    ('stocks',                               'Stocks [F]',                                      'section'),
    ('marchandises',                         'Marchandises',                                    'normal'),
    ('matieres fournitures consommables',    'Matières et fournitures consommables',            'normal'),
    ('produits cours',                       'Produits en cours',                               'normal'),
    ('produits intermediaires residuels',    'Produits intermédiaires et produits résiduels',   'normal'),
    ('produits finis',                       'Produits finis',                                  'normal'),
    ('creances actif circulant',             "Créances de l'actif circulant [G]",               'section'),
    ('fournisseurs debiteurs avances',       'Fournisseurs débiteurs, avances et acomptes',     'normal'),
    ('clients comptes rattaches',            'Clients et comptes rattachés',                    'normal'),
    ('personnel actif circulant',            'Personnel',                                       'normal'),
    ('etat actif',                           'État',                                            'normal'),
    ('comptes associes actif',               "Comptes d'associés",                              'normal'),
    ('autres debiteurs',                     'Autres débiteurs',                                'normal'),
    ('comptes regularisation actif',         'Comptes de régularisation - Actif',               'normal'),
    ('titres valeurs placement',             'Titres et valeurs de placement [H]',              'section'),
    ('ecarts conversion actif circulant',    'Écarts de conversion actif [I]',                  'section'),
    ('total ii actif',                       'TOTAL II (F+G+H+I)',                              'total'),
    ('tresorerie actif',                     'Trésorerie - Actif',                              'section'),
    ('cheques valeurs encaisser',            'Chèques et valeurs à encaisser',                  'normal'),
    ('banques tg ccp',                       'Banques, T.G et C.C.P',                          'normal'),
    ('caisse regie avances',                 "Caisse, Régie d'avances et accréditifs",         'normal'),
    ('total iii actif',                      'TOTAL III',                                       'total'),
    ('total general actif',                  'TOTAL GÉNÉRAL I+II+III',                          'total'),
]

PASSIF = [
    ('capitaux propres',                     'CAPITAUX PROPRES',                                'section'),
    ('capital social personnel',             'Capital social ou personnel (1)',                 'normal'),
    ('moins actionnaires capital',           'Moins : actionnaires, capital souscrit non appelé','normal'),
    ('capital appele',                       'Capital appelé',                                  'normal'),
    ('dont verse',                           'Dont versé',                                      'normal'),
    ('prime emission fusion',                "Prime d'émission, de fusion, d'apport",           'normal'),
    ('ecarts reevaluation',                  'Écarts de réévaluation',                          'normal'),
    ('reserve legale',                       'Réserve légale',                                  'normal'),
    ('autres reserves',                      'Autres réserves',                                 'normal'),
    ('report nouveau',                       'Report à nouveau (2)',                            'normal'),
    ('resultat instance affectation',        "Résultat en instance d'affectation (2)",          'normal'),
    ('resultat net exercice',                "Résultat net de l'exercice (2)",                  'normal'),
    ('total capitaux propres',               'Total des capitaux propres (A)',                  'total'),
    ('capitaux propres assimiles',           'Capitaux propres assimilés (B)',                  'section'),
    ('subvention investissement',            "Subvention d'investissement",                     'normal'),
    ('provisions reglementees',              'Provisions réglementées',                         'normal'),
    ('dettes financement',                   'Dettes de financement (C)',                       'section'),
    ('emprunts obligataires',                'Emprunts obligataires',                           'normal'),
    ('autres dettes financement',            'Autres dettes de financement',                    'normal'),
    ('provisions durables risques',          'Provisions durables pour risques et charges (D)', 'section'),
    ('provisions risques',                   'Provisions pour risques',                         'normal'),
    ('provisions charges',                   'Provisions pour charges',                         'normal'),
    ('ecarts conversion passif financement', 'Écarts de conversion passif (E)',                 'section'),
    ('augmentation creances immobilisees',   'Augmentation des créances immobilisées',          'normal'),
    ('diminution dettes financement',        'Diminution des dettes de financement',            'normal'),
    ('total i passif',                       'TOTAL I (A+B+C+D+E)',                             'total'),
    ('dettes passif circulant',              'Dettes du passif circulant (F)',                  'section'),
    ('fournisseurs comptes rattaches',       'Fournisseurs et comptes rattachés',               'normal'),
    ('clients crediteurs avances',           'Clients créditeurs, avances et acomptes',         'normal'),
    ('personnel passif',                     'Personnel',                                       'normal'),
    ('organismes sociaux',                   'Organismes sociaux',                              'normal'),
    ('etat passif',                          'État',                                            'normal'),
    ('comptes associes passif',              "Comptes d'associés",                              'normal'),
    ('autres creanciers',                    'Autres créanciers',                               'normal'),
    ('comptes regularisation passif',        'Comptes de régularisation passif',                'normal'),
    ('autres provisions risques charges',    'Autres provisions pour risques et charges (G)',   'normal'),
    ('ecarts conversion passif circulant',   'Écarts de conversion passif (H)',                 'normal'),
    ('total ii passif',                      'TOTAL II (F+G+H)',                                'total'),
    ('tresorerie passif',                    'TRÉSORERIE PASSIF',                               'section'),
    ('credits escompte',                     "Crédits d'escompte",                              'normal'),
    ('credits tresorerie',                   'Crédits de trésorerie',                           'normal'),
    ('banques soldes crediteurs',            'Banques (soldes créditeurs)',                     'normal'),
    ('total iii passif',                     'TOTAL III',                                       'total'),
    ('total general passif',                 'TOTAL GÉNÉRAL I+II+III',                          'total'),
]

CPC = [
    ('produits exploitation',                "PRODUITS D'EXPLOITATION",                         'section'),
    ('ventes marchandises',                  "Ventes de marchandises (en l'état)",             'normal'),
    ('ventes biens services',                'Ventes de biens et services produits',            'normal'),
    ('chiffres affaires',                    "Chiffre d'affaires",                              'normal'),
    ('variation stocks produits',            'Variation de stocks de produits (±)',             'normal'),
    ('immobilisations produites',            "Immobilisations produites par l'entreprise",      'normal'),
    ('subventions exploitation',             "Subventions d'exploitation",                      'normal'),
    ('autres produits exploitation',         "Autres produits d'exploitation",                  'normal'),
    ('reprises exploitation',                "Reprises d'exploitation ; transferts de charges", 'normal'),
    ('total i cpc',                          'Total I',                                         'total'),
    ('charges exploitation',                 "CHARGES D'EXPLOITATION",                          'section'),
    ('achats revendus marchandises',         'Achats revendus de marchandises',                 'normal'),
    ('achats consommes matieres',            'Achats consommés de matières et fournitures',     'normal'),
    ('autres charges externes',              'Autres charges externes',                         'normal'),
    ('impots taxes',                         'Impôts et taxes',                                 'normal'),
    ('charges personnel',                    'Charges de personnel',                            'normal'),
    ('autres charges exploitation',          "Autres charges d'exploitation",                   'normal'),
    ('dotations exploitation',               "Dotations d'exploitation",                        'normal'),
    ('total ii cpc',                         'Total II',                                        'total'),
    ('resultat exploitation',                "RÉSULTAT D'EXPLOITATION (I-II)",                  'result'),
    ('produits financiers',                  'PRODUITS FINANCIERS',                             'section'),
    ('produits titres participation',        'Produits des titres de participation',            'normal'),
    ('gains change',                         'Gains de change',                                 'normal'),
    ('interets autres produits financiers',  'Intérêts et autres produits financiers',          'normal'),
    ('reprises financieres',                 'Reprises financières ; transferts de charges',    'normal'),
    ('total iv',                             'Total IV',                                        'total'),
    ('charges financieres section',          'CHARGES FINANCIÈRES',                             'section'),
    ('charges interets',                     "Charges d'intérêts",                              'normal'),
    ('pertes change',                        'Pertes de change',                                'normal'),
    ('autres charges financieres',           'Autres charges financières',                      'normal'),
    ('dotations financieres',                'Dotations financières',                           'normal'),
    ('total v',                              'Total V',                                         'total'),
    ('resultat financier',                   'RÉSULTAT FINANCIER (IV-V)',                       'result'),
    ('resultat courant',                     'RÉSULTAT COURANT (III+VI)',                       'result'),
    ('produits non courants',                'PRODUITS NON COURANTS',                           'section'),
    ('produits cessions immobilisations',    "Produits des cessions d'immobilisations",         'normal'),
    ('subventions equilibre',                "Subventions d'équilibre",                         'normal'),
    ('reprises subventions investissement',  "Reprises sur subventions d'investissement",       'normal'),
    ('autres produits non courants',         'Autres produits non courants',                    'normal'),
    ('reprises non courantes',               'Reprises non courantes ; transferts de charges',  'normal'),
    ('total viii',                           'Total VIII',                                      'total'),
    ('charges non courantes section',        'CHARGES NON COURANTES',                           'section'),
    ('valeurs nettes amortissements',        "Valeurs nettes d'amortissement des immob. cédées",'normal'),
    ('subventions accordees',                'Subventions accordées',                           'normal'),
    ('autres charges non courantes',         'Autres charges non courantes',                    'normal'),
    ('dotations non courantes',              'Dotations non courantes',                         'normal'),
    ('total ix',                             'Total IX',                                        'total'),
    ('resultat non courant',                 'RÉSULTAT NON COURANT (VIII-IX)',                  'result'),
    ('resultat avant impots',                'RÉSULTAT AVANT IMPÔTS (VII+X)',                  'result'),
    ('impots resultats',                     'IMPÔTS SUR LES RÉSULTATS',                       'normal'),
    ('resultat net',                         'RÉSULTAT NET (XI-XII)',                           'result'),
    ('total produits',                       'TOTAL DES PRODUITS (I+IV+VIII)',                  'total'),
    ('total charges',                        'TOTAL DES CHARGES (II+V+IX+XII)',                 'total'),
    ('resultat net total',                   'RÉSULTAT NET (Total produits - Total charges)',   'result'),
]

# ══════════════════════════════════════════════════════════════════
# NORMALISATION
# ══════════════════════════════════════════════════════════════════

def _norm(s: str) -> str:
    s = unicodedata.normalize('NFD', str(s))
    s = s.encode('ascii', 'ignore').decode().lower()
    s = re.sub(r'^\s*[\*\.]+\s*', '', s)
    s = re.sub(r'^(xvi|xiv|xiii|xii|xi|ix|viii|vii|vi|iv|v|iii|ii|i)\s*[\.\-\=]?\s*', '', s, flags=re.I)
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'chiffres?', 'chiffre', s)
    s = re.sub(r'reserves?', 'reserve', s)
    s = re.sub(r'reprises?', 'reprise', s)
    s = re.sub(r'provisions?', 'provision', s)
    s = re.sub(r'subventions?', 'subvention', s)
    return re.sub(r'\s+', ' ', s).strip()

def _parse(s) -> float | None:
    if s is None: return None
    if isinstance(s, (int, float)): return float(s)
    s = str(s).strip().replace('\xa0','').replace(' ','')
    if not s or s in ['-','—']: return None
    neg = s.startswith('-'); s = s.lstrip('-')
    m = re.match(r'^(\d{1,3}(?:\.\d{3})*),(\d{2})$', s)
    if m: s = m.group(1).replace('.','') + '.' + m.group(2)
    elif re.match(r'^\d+,\d{2}$', s): s = s.replace(',','.')
    elif re.match(r'^\d+(\.\d+)?$', s): pass
    else: return None
    try: return -float(s) if neg else float(s)
    except: return None

# ══════════════════════════════════════════════════════════════════
# MATCHING
# ══════════════════════════════════════════════════════════════════

def match_label(label: str, template: list, used: set = None) -> int:
    if used is None: used = set()
    n = _norm(label)

    _romain_seul = {
        'xi':   'resultat avant impots',
        'xii':  'impots resultats',
        'xiii': 'resultat net',
        'xiv':  'total produits',
        'xv':   'total charges',
        'xvi':  'resultat net total',
    }
    _label_brut = re.sub(r'[^a-zA-Z]', '', str(label).lower().strip())
    if _label_brut in _romain_seul:
        def _ff_early(key):
            for i, (k, _, _) in enumerate(template):
                if k == key:
                    return i if i not in used else -1
            return None
        r = _ff_early(_romain_seul[_label_brut])
        if r is not None: return r

    if not n or len(n) < 2: return -1

    syn_idx = lookup_in_template(label, template)
    if syn_idx >= 0 and syn_idx not in used:
        return syn_idx

    def _first_free(key):
        for i, (k, _, _) in enumerate(template):
            if k == key:
                return i if i not in used else -1
        return None

    if 'total' in n:
        _total_rules = [
            ('general',                  ['total general actif', 'total general passif']),
            ('i ii iii',                 ['total general actif', 'total general passif']),
            ('l ii iii',                 ['total general actif', 'total general passif']),
            ('iii',                      ['total iii actif', 'total iii passif']),
            ('f g h i',                  ['total ii actif']),
            ('f g h',                    ['total ii passif']),
            ('a b c d e',                ['total i actif', 'total i passif']),
            ('total des produits',       ['total produits']),
            ('total des charges',        ['total charges']),
            ('total viii',               ['total viii']),
            ('total ix',                 ['total ix']),
            ('total iv',                 ['total iv']),
            ('total ii',                 ['total ii cpc', 'total ii actif', 'total ii passif']),
            ('total v',                  ['total v']),
            ('total i',                  ['total i cpc', 'total i actif', 'total i passif']),
        ]
        for pat, keys in _total_rules:
            if pat in n:
                for key in keys:
                    r = _first_free(key)
                    if r is not None:
                        return r
                break

    if 'ecart' in n and 'conversion' in n:
        if 'passif' not in n:
            if 'circulant' in n or 'element' in n:
                r = _first_free('ecarts conversion actif circulant')
                if r is not None: return r
            r = _first_free('ecarts conversion actif immobilise')
            if r is not None: return r
            return _first_free('ecarts conversion actif circulant')

    if 'ecart' in n and 'conversion' in n and 'passif' in n:
        if 'circulant' in n or 'element' in n:
            return _first_free('ecarts conversion passif circulant')
        else:
            return _first_free('ecarts conversion passif financement')

    if 'reprise' in n:
        if 'non courant' in n:
            return _first_free('reprises non courantes')
        if 'financier' in n:
            return _first_free('reprises financieres')
        if 'exploitation' in n:
            return _first_free('reprises exploitation')

    if 'dotation' in n:
        if 'non courant' in n:
            return _first_free('dotations non courantes')
        if 'financier' in n:
            return _first_free('dotations financieres')
        if 'exploitation' in n:
            return _first_free('dotations exploitation')

    if 'autres' in n and 'charge' in n:
        if 'non courant' in n:
            return _first_free('autres charges non courantes')
        if 'financier' in n:
            return _first_free('autres charges financieres')
        if 'exploitation' in n:
            return _first_free('autres charges exploitation')
        if 'externe' in n:
            return _first_free('autres charges externes')

    if 'autre' in n and 'produit' in n:
        if 'non courant' in n:
            return _first_free('autres produits non courants')
        if 'exploitation' in n:
            return _first_free('autres produits exploitation')

    if 'resultat' in n:
        _results = [
            ('resultat net xi',          'resultat net'),
            ('resultat net total',       'resultat net total'),
            ('resultat avant impot',     'resultat avant impots'),
            ('resultat non courant',     'resultat non courant'),
            ('resultat courant',         'resultat courant'),
            ('resultat financier',       'resultat financier'),
            ('resultat exploitation',    'resultat exploitation'),
        ]
        for pat, key in _results:
            if n.startswith(pat) or pat in n:
                return _first_free(key)

    words_label = set(w for w in n.split() if len(w) > 2)
    if not words_label: return -1

    scores = []
    for i, (key, display, typ) in enumerate(template):
        if i in used: continue
        n_key = _norm(key)
        words_key = set(w for w in n_key.split() if len(w) > 2)
        if not words_key: continue

        common = len(words_label & words_key)
        union  = len(words_label | words_key)
        score  = common / union if union > 0 else 0

        if n.startswith(n_key[:8]) or n_key.startswith(n[:8]):
            score = min(score + 0.2, 1.0)

        if score >= 0.35:
            scores.append((score, i))

    if not scores: return -1
    scores.sort(reverse=True)
    return scores[0][1]

# ══════════════════════════════════════════════════════════════════
# EXTRACTION
# ══════════════════════════════════════════════════════════════════

def _is_rotated(v) -> bool:
    if not v: return False
    parts = [p.strip() for p in str(v).split('\n') if p.strip()]
    return len(parts) >= 3 and all(len(p) <= 2 and p.replace('.','').isalpha() for p in parts)

def _is_index_row(row) -> bool:
    vals = [str(c).strip() for c in (row or []) if c is not None]
    if len(vals) < 2: return False
    try:
        nums = [int(v) for v in vals]
        return nums == list(range(len(nums)))
    except: return False

def _has_fused(table) -> bool:
    count = 0
    for row in (table or []):
        for cell in (row or []):
            if cell and '\n' in str(cell):
                parts = [x for x in str(cell).split('\n') if _parse(x) is not None]
                if len(parts) > 1: count += 1
    return count >= 3

def _detect_val_cols(table) -> list:
    """
    Détecte les colonnes contenant des valeurs numériques.

    FIX : les 4 dernières lignes du CPC (XIV TOTAL DES PRODUITS,
    XV TOTAL DES CHARGES, XVI RÉSULTAT NET) sont dans un tableau
    séparé de seulement 3 lignes sur la page 5. L'ancienne version
    faisait [2:] sur ce tableau → 1 seule ligne analysée → aucune
    colonne n'atteignait le seuil de 2 → val_cols = [] → lignes ignorées.

    Correction : adapter le skip et le seuil minimum à la taille du tableau.
    - Tableaux >= 5 lignes : skip 2 lignes de header, seuil 2 occurrences
    - Petits tableaux < 5 lignes : skip 0, seuil 1 occurrence
    """
    hits = {}
    tbl = table or []
    skip     = 2 if len(tbl) >= 5 else 0
    min_hits = 2 if len(tbl) >= 5 else 1
    for row in tbl[skip:]:
        if not row: continue
        for ci, c in enumerate(row):
            if not c: continue
            s = str(c).strip().replace('\xa0','').replace(' ','')
            if re.match(r'^-?\d{1,3}(\.\d{3})*,\d{2}$', s) or re.match(r'^-?\d+,\d{2}$', s):
                hits[ci+1] = hits.get(ci+1, 0) + 1
    return sorted(k for k, v in hits.items() if v >= min_hits)

def _xy_rows(page) -> list:
    """Extraction X/Y pour pages avec cellules fusionnées."""
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words: return []

    def is_num(t):
        return (bool(re.match(r'^-?\d+$', t.replace(',','').replace('.','')))
                and len(t.replace(',','').replace('.','')) >= 1)

    num_ws = [w for w in words if is_num(w['text']) and w['x0'] > 100]
    if not num_ws: return []
    thresh = min(w['x0'] for w in num_ws) - 5

    lines = defaultdict(list)
    for w in words:
        lines[round(w['top'] / 3) * 3].append(w)

    result = []
    for y in sorted(lines):
        row = sorted(lines[y], key=lambda w: w['x0'])
        lw = [w for w in row if w['x0'] < thresh]
        nw = [w for w in row if w['x0'] >= thresh and is_num(w['text'])]

        filt = [w for w in lw if not (len(w['text'])<=1
                                       and re.match(r'^[A-Z.]$', w['text'])
                                       and w['x0'] < 50)]
        label = ''
        if filt:
            label = filt[0]['text']
            for i in range(1, len(filt)):
                gap = filt[i]['x0'] - filt[i-1]['x1']
                label += filt[i]['text'] if gap <= 1 else ' ' + filt[i]['text']
            label = re.sub(r' +', ' ', label).strip()

        vals = []
        if nw:
            grp = [nw[0]]
            for w in nw[1:]:
                if w['x0'] - grp[-1]['x1'] < 18: grp.append(w)
                else:
                    v = _parse(''.join(x['text'] for x in grp))
                    if v is not None: vals.append(v)
                    grp = [w]
            v = _parse(''.join(x['text'] for x in grp))
            if v is not None: vals.append(v)

        if label and vals:
            result.append((label, vals))
        elif label and not vals:
            result.append((label, []))
        elif not label and vals and result:
            last_label, last_vals = result[-1]
            if not last_vals:
                result[-1] = (last_label, vals)

    return [(l, v) for l, v in result if v]

def _extract_section(pdf, page_indices: list, is_actif: bool = False) -> list[tuple[str, list]]:
    """
    Extrait les paires (label, [valeurs]) d'une section.
    Choisit automatiquement X/Y ou extract_tables().
    """
    all_rows = []
    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        page = pdf.pages[idx]
        tables = page.extract_tables()
        good = [t for t in tables
                if t and len(t) >= 3 and t[0] and len(t[0]) >= 2
                and sum(1 for r in t[1:] if any(c for c in r if c)) >= 2]
        if not good: continue

        if any(_has_fused(t) for t in good):
            rows = _xy_rows(page)
        else:
            rows = []
            for t in good:
                val_cols = _detect_val_cols(t)
                for row in t:
                    if not row or _is_index_row(row): continue
                    if len(row) < 2: continue
                    col1 = str(row[1]).strip().replace('\n',' ') if len(row)>1 and row[1] else ''
                    if _ROMAIN_RE.match(col1.strip()):
                        label = str(row[2]).strip().replace('\n',' ') if len(row)>2 and row[2] else ''
                    else:
                        label = col1
                    if _is_rotated(label): label = ''
                    if not label or len(label) < 2: continue
                    if _parse(label) is not None: continue

                    def gv(col):
                        if col and len(row) >= col: return _parse(row[col-1])
                        return None
                    nv = len(val_cols)
                    if is_actif and nv == 3:
                        gap = val_cols[1] - val_cols[0]
                        if gap > 1:
                            vals = [gv(val_cols[0]), None, gv(val_cols[1]), gv(val_cols[2])]
                        else:
                            vals = [gv(val_cols[0]), gv(val_cols[1]), gv(val_cols[2]), None]
                    else:
                        vals = [gv(val_cols[i]) for i in range(min(nv, 4))]
                    if vals or label:
                        rows.append((label, vals))

        all_rows.extend(rows)
    return all_rows

def _build_value_map(rows: list, template: list) -> dict:
    """
    Mappe les lignes extraites vers les indices du template.
    Retourne {template_idx: [v1, v2, v3, v4]}
    """
    result = {}
    used = set()

    for label, vals in rows:
        idx = match_label(label, template, used=used)
        if idx < 0: continue
        if idx in result: continue
        result[idx] = vals
        used.add(idx)

    return result

# ══════════════════════════════════════════════════════════════════
# INFOS
# ══════════════════════════════════════════════════════════════════

def extract_info(pdf) -> dict:
    info = {}
    for i in range(min(2, len(pdf.pages))):
        text = pdf.pages[i].extract_text() or ''
        for key, pat in [
            ('raison_sociale',      r'[Rr]aison\s+[Ss]ociale\s*:?\s*([A-Z][^\n]{3,60})'),
            ('identifiant_fiscal',  r'[Ii]dentifiant\s+[Ff]iscal\s*:?\s*(\d+)'),
            ('taxe_professionnelle',r'[Tt]axe\s+[Pp]rof\w*\.?\s*:?\s*([\d\s]+)'),
            ('adresse',             r'[Aa]dresse\s*:?\s*([^\n]{5,60})'),
        ]:
            if key not in info:
                m = re.search(pat, text, re.IGNORECASE)
                if m: info[key] = m.group(1).strip()

        if 'exercice' not in info:
            for pat in [
                r'[Ee]xercice\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})',
                r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})',
            ]:
                m = re.search(pat, text)
                if m:
                    info['exercice']     = f"Du {m.group(1)} au {m.group(2)}"
                    info['exercice_fin'] = m.group(2)
                    break

    for k in ('raison_sociale','identifiant_fiscal','taxe_professionnelle',
              'adresse','exercice','exercice_fin'):
        info.setdefault(k, '')
    return info

# ══════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════

def parse(pdf_path: str) -> dict:
    """
    Parse un PDF AMMC (5 pages).
    Retourne {info, actif, passif, cpc} où chaque section est
    {template_idx: [v1, v2, v3, v4]}.
    """
    pdf = pdfplumber.open(pdf_path)
    info = extract_info(pdf)

    logger.info(f"AMMC parser : {info.get('raison_sociale','?')} | {info.get('exercice','?')}")

    actif_rows  = _extract_section(pdf, [1], is_actif=True)
    passif_rows = _extract_section(pdf, [2])
    cpc_rows    = _extract_section(pdf, [3, 4])
    pdf.close()

    actif_map  = _build_value_map(actif_rows,  ACTIF)
    passif_map = _build_value_map(passif_rows, PASSIF)
    cpc_map    = _build_value_map(cpc_rows,    CPC)

    logger.info(f"  Actif  : {len(actif_map)}/{len(ACTIF)} postes")
    logger.info(f"  Passif : {len(passif_map)}/{len(PASSIF)} postes")
    logger.info(f"  CPC    : {len(cpc_map)}/{len(CPC)} postes")

    return {
        'info':   info,
        'actif':  actif_map,
        'passif': passif_map,
        'cpc':    cpc_map,
        'format': 'AMMC',
        'templates': {'actif': ACTIF, 'passif': PASSIF, 'cpc': CPC},
    }
