"""
core/pdf_to_excel.py
Extraction PDF fiscal → Excel structuré.

Deux pistes automatiques :
  - AMMC (5 pages) : pages 2=Actif, 3=Passif, 4-5=CPC
  - DGI  (7 pages) : pages 2-3=Actif, 4=Passif, 5-7=CPC

Pour chaque piste :
  - Si le tableau a des cellules fusionnées → extraction X/Y (mots par position)
  - Sinon → extract_tables() direct
  - Même structure Excel finale pour les deux formats
"""

import re
from collections import defaultdict
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from utils.logger import get_logger

logger = get_logger(__name__)

# ═══════════════════════════════════════════════════════════
# STYLES
# ═══════════════════════════════════════════════════════════

C_DARK   = "1F3864"
C_MED    = "2E75B6"
C_LIGHT  = "D6E4F0"
C_WHITE  = "FFFFFF"
C_GRAY   = "F5F7FA"
C_BORDER = "B8CCE4"
NUM_FMT  = '#,##0.00;[Red]-#,##0.00;"-"'

def _side():
    return Side(style='thin', color=C_BORDER)

def _border():
    s = _side()
    return Border(top=s, bottom=s, left=s, right=s)

def _c(ws, r, c, value=None, bg=C_WHITE, fg="222222", bold=False,
       align="left", num_fmt=None, size=9, wrap=True, indent=0):
    cell = ws.cell(row=r, column=c, value=value)
    cell.font      = Font(name="Arial", size=size, bold=bold, color=fg)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=wrap, indent=indent)
    cell.border    = _border()
    if num_fmt:
        cell.number_format = num_fmt
    return cell

def _row_style(label: str):
    """(bg, fg, bold) selon le type de ligne."""
    n = label.strip().upper()
    if re.match(r'^TOTAL', n):
        return C_DARK, C_WHITE, True
    if re.match(r'^R[EÉ]SULTAT', n):
        return "2E4057", C_WHITE, True
    # Ligne de section (tout majuscules > 5 car)
    alphanum = re.sub(r'[^A-Z0-9]', '', n)
    if len(alphanum) > 5 and alphanum == re.sub(r'[^A-Z0-9]', '', label.upper()):
        return C_LIGHT, C_DARK, True
    return C_WHITE, "333333", False

# ═══════════════════════════════════════════════════════════
# UTILITAIRES NETTOYAGE
# ═══════════════════════════════════════════════════════════

def _is_rotated(v) -> bool:
    """Lettres rotatives A\nC\nT\nI\nF → col décorative."""
    if not v: return False
    parts = [p.strip() for p in str(v).split('\n') if p.strip()]
    return len(parts) >= 3 and all(len(p) <= 2 and p.replace('.','').isalpha() for p in parts)

def _is_index_row(row: list) -> bool:
    """Ligne parasite 0 1 2 3 4 ..."""
    vals = [str(c).strip() for c in row if c is not None]
    if len(vals) < 2: return False
    try:
        nums = [int(v) for v in vals]
        return nums == list(range(len(nums)))
    except (ValueError, TypeError):
        return False

def _clean(v) -> str:
    """Nettoyage de base d'une cellule."""
    if not v: return ''
    s = str(v).strip()
    if _is_rotated(s): return ''
    s = re.sub(r'\n+', ' ', s)
    return re.sub(r' {2,}', ' ', s).strip()

def _parse(s) -> float | None:
    """Parse nombre FR : '1 234 567,89' → 1234567.89"""
    if not s: return None
    s = str(s).strip().replace('\xa0', '').replace(' ', '')
    if not s or s in ['-', '—', '/']: return None
    neg = s.startswith('-')
    s = s.lstrip('-')
    m = re.match(r'^(\d{1,3}(?:\.\d{3})*),(\d{2})$', s)
    if m:
        s = m.group(1).replace('.', '') + '.' + m.group(2)
    elif re.match(r'^\d+,\d{2}$', s):
        s = s.replace(',', '.')
    elif re.match(r'^\d+$', s):
        pass
    else:
        return None
    try:
        return -float(s) if neg else float(s)
    except ValueError:
        return None

def _is_num_tok(t: str) -> bool:
    """Token numérique pour extraction X/Y."""
    return (bool(re.match(r'^-?\d+$', t.replace(',', '').replace('.', '')))
            and len(t.replace(',', '').replace('.', '')) >= 1)

# ═══════════════════════════════════════════════════════════
# EXTRACTION X/Y (fallback cellules fusionnées)
# ═══════════════════════════════════════════════════════════

def _extract_xy(page) -> list[tuple[str, list]]:
    """
    Extrait les lignes d'une page via position X/Y des mots.
    Retourne [(label, [v1, v2, v3, v4]), ...]
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words: return []

    num_ws = [w for w in words if _is_num_tok(w['text']) and w['x0'] > 100]
    if not num_ws: return []
    thresh = min(w['x0'] for w in num_ws) - 5

    lines = defaultdict(list)
    for w in words:
        lines[round(w['top'] / 3) * 3].append(w)

    result = []
    for y in sorted(lines):
        row = sorted(lines[y], key=lambda w: w['x0'])
        lw = [w for w in row if w['x0'] < thresh]
        nw = [w for w in row if w['x0'] >= thresh and _is_num_tok(w['text'])]

        # Reconstruire le label (filtrer lettres rotatives x<50 longueur=1)
        filtered = [w for w in lw
                    if not (len(w['text']) <= 1
                            and re.match(r'^[A-Z.]$', w['text'])
                            and w['x0'] < 50)]
        label = ''
        if filtered:
            label = filtered[0]['text']
            for i in range(1, len(filtered)):
                gap = filtered[i]['x0'] - filtered[i-1]['x1']
                label += filtered[i]['text'] if gap <= 1 else ' ' + filtered[i]['text']
            label = re.sub(r' {2,}', ' ', label).strip()

        # Fusionner tokens numériques adjacents → valeurs
        vals = []
        if nw:
            grp = [nw[0]]
            for w in nw[1:]:
                if w['x0'] - grp[-1]['x1'] < 18:
                    grp.append(w)
                else:
                    v = _parse(''.join(x['text'] for x in grp))
                    if v is not None: vals.append(v)
                    grp = [w]
            v = _parse(''.join(x['text'] for x in grp))
            if v is not None: vals.append(v)

        # Label sans valeur → garder pour rattachement ultérieur
        if label:
            result.append((label, vals))
        elif vals and result:
            # Valeurs orphelines → rattacher au dernier label vide
            last_label, last_vals = result[-1]
            if not last_vals:
                result[-1] = (last_label, vals)

    # Ne garder que les lignes avec valeurs
    return [(l, v) for l, v in result if v]

# ═══════════════════════════════════════════════════════════
# EXTRACTION TABLEAU STANDARD
# ═══════════════════════════════════════════════════════════

def _has_fused(table) -> bool:
    """Détecte cellules avec plusieurs valeurs (\n)."""
    count = 0
    for row in (table or []):
        for cell in (row or []):
            if cell and '\n' in str(cell):
                parts = [x for x in str(cell).split('\n') if _parse(x) is not None]
                if len(parts) > 1:
                    count += 1
    return count >= 3

def _extract_table(table) -> list[tuple[str, list]]:
    """
    Extrait les lignes d'un tableau standard.
    Retourne [(label, [v1, v2, ...]), ...]
    """
    rows = []
    for row in table:
        if not row: continue
        if _is_index_row(row): continue

        cells = [_clean(c) for c in row]

        # Label = première cellule non-numérique de longueur > 1
        label = ''
        label_col = -1
        for ci, c in enumerate(cells[:3]):
            if c and len(c) > 1 and _parse(c) is None:
                label = c
                label_col = ci
                break
        if not label or len(label) < 2:
            continue

        vals = []
        for ci, c in enumerate(cells):
            if ci == label_col: continue
            v = _parse(c)
            if v is not None:
                vals.append(v)

        rows.append((label, vals))
    return rows

def _get_rows(pdf, page_indices: list) -> list[tuple[str, list]]:
    """
    Extrait toutes les lignes des pages indiquées.
    Choisit automatiquement X/Y ou tableau selon le contenu.
    """
    all_rows = []
    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        page = pdf.pages[idx]
        tables = page.extract_tables()
        good = [t for t in tables
                if t and len(t) >= 3 and t[0] and len(t[0]) >= 2
                and sum(1 for r in t[2:] if any(c for c in r if c)) >= 3]

        if not good:
            continue

        if any(_has_fused(t) for t in good):
            logger.info(f"  page {idx+1} → X/Y (cellules fusionnées)")
            rows = _extract_xy(page)
        else:
            logger.info(f"  page {idx+1} → extract_tables()")
            rows = []
            for t in good:
                rows.extend(_extract_table(t))

        all_rows.extend(rows)

    # Dédoublonner : même label normalisé → garder le premier
    seen = set()
    deduped = []
    for label, vals in all_rows:
        key = re.sub(r'\W', '', label.lower())
        if key not in seen:
            seen.add(key)
            deduped.append((label, vals))

    return deduped

# ═══════════════════════════════════════════════════════════
# INFOS GÉNÉRALES
# ═══════════════════════════════════════════════════════════

def _info(pdf) -> dict:
    d = {}
    for i in range(min(2, len(pdf.pages))):
        text = pdf.pages[i].extract_text() or ''
        for key, pat in [
            ('raison_sociale',      r'[Rr]aison\s+[Ss]ociale\s*:?\s*([A-Z][^\n]{3,60})'),
            ('identifiant_fiscal',  r'[Ii]dentifiant\s+[Ff]iscal\s*:?\s*(\d+)'),
            ('taxe_professionnelle',r'[Tt]axe\s+[Pp]rof\w*\.?\s*:?\s*([\d\s]+)'),
            ('adresse',             r'[Aa]dresse\s*:?\s*([^\n]{5,60})'),
        ]:
            if key not in d:
                m = re.search(pat, text, re.IGNORECASE)
                if m: d[key] = m.group(1).strip()

        if 'exercice' not in d:
            for pat in [
                r'p[eé]riode\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})',
                r'[Ee]xercice\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})',
                r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})',
            ]:
                m = re.search(pat, text)
                if m:
                    d['exercice']     = f"Du {m.group(1)} au {m.group(2)}"
                    d['exercice_fin'] = m.group(2)
                    break

    for k in ('raison_sociale', 'identifiant_fiscal', 'taxe_professionnelle',
              'adresse', 'exercice', 'exercice_fin'):
        d.setdefault(k, '')
    return d

# ═══════════════════════════════════════════════════════════
# ÉCRITURE EXCEL
# ═══════════════════════════════════════════════════════════

def _header_block(ws, title: str, info: dict, n_cols: int) -> int:
    """Bloc titre + infos société. Retourne la prochaine ligne libre."""
    raison   = info.get('raison_sociale', '')
    if_num   = info.get('identifiant_fiscal', '')
    exercice = info.get('exercice', '')

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    _c(ws, 1, 1, title, bg=C_DARK, fg=C_WHITE, bold=True, align='center', size=12)
    ws.row_dimensions[1].height = 22

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols-1)
    _c(ws, 2, 1, f"Raison sociale : {raison}", bg=C_LIGHT, fg=C_DARK,
       bold=True, size=9, indent=1)
    _c(ws, 2, n_cols, f"IF : {if_num}", bg=C_LIGHT, fg=C_DARK,
       align='right', size=9)

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=n_cols)
    _c(ws, 3, 1, f"Exercice : {exercice}", bg=C_GRAY, fg="555555", size=9, indent=1)

    ws.row_dimensions[4].height = 4   # séparateur vide
    return 5

def _col_headers(ws, r: int, cols: list) -> None:
    """cols = [(titre, largeur), ...]"""
    for ci, (h, w) in enumerate(cols, 1):
        _c(ws, r, ci, h, bg=C_MED, fg=C_WHITE, bold=True,
           align='center', size=9, wrap=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[r].height = 30
    ws.freeze_panes = ws.cell(r + 1, 1)

def _data_rows(ws, r: int, rows: list, n_vals: int,
               label_col: int = 1) -> int:
    """
    Écrit les lignes de données.
    label_col : colonne du label (1-based)
    Les valeurs commencent à label_col+1.
    """
    written = 0
    for label, vals in rows:
        bg, fg, bold = _row_style(label)
        ws.row_dimensions[r].height = 15 if bg == C_WHITE else 17
        indent = 1 if bg == C_WHITE else 0

        _c(ws, r, label_col, label, bg=bg, fg=fg, bold=bold,
           align='left', size=9, indent=indent)

        for ci in range(n_vals):
            v = vals[ci] if ci < len(vals) else None
            _c(ws, r, label_col + ci + 1, 0 if v is None else v,
               bg=bg, fg=fg, bold=bold, align='right', size=9, num_fmt=NUM_FMT)

        r += 1
        written += 1
    return written

# ═══════════════════════════════════════════════════════════
# FEUILLES
# ═══════════════════════════════════════════════════════════

def _sheet_ident(wb, info: dict):
    ws = wb.create_sheet("1 - Identification")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 52

    ws.merge_cells('A1:B1')
    _c(ws, 1, 1, "PIÈCES ANNEXES À LA DÉCLARATION FISCALE",
       bg=C_DARK, fg=C_WHITE, bold=True, align='center', size=13)
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:B2')
    _c(ws, 2, 1, "IMPÔTS SUR LES SOCIÉTÉS — Modèle Comptable Normal (loi 9-88)",
       bg=C_MED, fg=C_WHITE, align='center', size=10)
    ws.row_dimensions[2].height = 18

    for i, (lbl, val) in enumerate([
        ("Raison sociale",        info.get('raison_sociale', '—')),
        ("Identifiant fiscal",    info.get('identifiant_fiscal', '—')),
        ("Taxe professionnelle",  info.get('taxe_professionnelle', '—')),
        ("Adresse",               info.get('adresse', '—')),
        ("Exercice",              info.get('exercice', '—')),
    ], 4):
        ws.row_dimensions[i].height = 18
        _c(ws, i, 1, lbl, bg=C_LIGHT, fg=C_DARK, bold=True, size=9, indent=1)
        _c(ws, i, 2, val, bg=C_WHITE,  fg="222222", size=9, indent=1)


def _sheet_actif(wb, info: dict, rows: list) -> int:
    """
    Bilan Actif : Label | Brut | Amort.&Prov. | Net(N) | Net(N-1)
    """
    ws = wb.create_sheet("2 - Bilan Actif")
    ws.sheet_view.showGridLines = False
    n = 5
    r = _header_block(ws, "BILAN ACTIF", info, n)
    _col_headers(ws, r, [
        ("DÉSIGNATION",          50),
        ("BRUT",                 18),
        ("AMORT. & PROV.",       18),
        ("NET — EXERCICE N",     18),
        ("NET — EXERCICE N-1",   18),
    ])
    return _data_rows(ws, r + 1, rows, n_vals=4)


def _sheet_passif(wb, info: dict, rows: list) -> int:
    """
    Bilan Passif : Label | Exercice N | Exercice N-1
    """
    ws = wb.create_sheet("3 - Bilan Passif")
    ws.sheet_view.showGridLines = False
    n = 3
    r = _header_block(ws, "BILAN PASSIF", info, n)
    _col_headers(ws, r, [
        ("DÉSIGNATION",     54),
        ("EXERCICE N",      20),
        ("EXERCICE N-1",    20),
    ])
    written = _data_rows(ws, r + 1, rows, n_vals=2)

    # Note légale
    r2 = r + 1 + written + 1
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=3)
    c = ws.cell(r2, 1)
    c.value = "(1) Capital personnel débiteur.  (2) Bénéficiaire (+) / Déficitaire (−)."
    c.font  = Font(name="Arial", italic=True, size=8, color="888888")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return written


def _sheet_cpc(wb, info: dict, rows: list) -> int:
    """
    CPC : N° | Label | PropresN | ExercPréc | TotalN | TotalN-1
    """
    ws = wb.create_sheet("4 - CPC")
    ws.sheet_view.showGridLines = False
    n = 6
    r = _header_block(ws, "COMPTE DE PRODUITS ET CHARGES (Hors Taxes)", info, n)
    _col_headers(ws, r, [
        ("N°",                     5),
        ("DÉSIGNATION",            44),
        ("PROPRES À\nL'EXERCICE",  18),
        ("EXERCICES\nPRÉCÉDENTS",  18),
        ("TOTAUX\nEXERCICE N",     18),
        ("TOTAUX\nEXERCICE N-1",   18),
    ])
    r += 1

    # CPC : on essaie de détecter le numéro romain en tête de label
    ROMAIN = re.compile(
        r'^(I{1,3}|IV|VI{0,3}|IX|XI{0,3}|XIV|XV|XVI)\b\.?\s*', re.I)

    written = 0
    for label, vals in rows:
        bg, fg, bold = _row_style(label)
        ws.row_dimensions[r].height = 15 if bg == C_WHITE else 17
        indent = 1 if bg == C_WHITE else 0

        # Extraire numéro romain
        m = ROMAIN.match(label)
        num = m.group(1).upper() if m else ''
        clean_label = ROMAIN.sub('', label).strip() if m else label

        _c(ws, r, 1, num or None, bg=bg,
           fg=fg if bg != C_WHITE else C_MED,
           bold=True, align='center', size=8)
        _c(ws, r, 2, clean_label, bg=bg, fg=fg, bold=bold,
           align='left', size=9, indent=indent)

        for ci in range(4):
            v = vals[ci] if ci < len(vals) else None
            _c(ws, r, ci + 3, 0 if v is None else v,
               bg=bg, fg=fg, bold=bold, align='right', size=9, num_fmt=NUM_FMT)

        r += 1
        written += 1

    # Notes
    r2 = r + 1
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=6)
    c = ws.cell(r2, 1)
    c.value = ("(1) Stock final − Stock initial.  "
               "(2) Achats revendus = Achats − Variation de stock.")
    c.font  = Font(name="Arial", italic=True, size=8, color="888888")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return written

# ═══════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ═══════════════════════════════════════════════════════════

def convert(pdf_path: str, output_path: str) -> dict:
    """
    Convertit un PDF fiscal → Excel structuré.
    Détecte automatiquement AMMC (5 pages) ou DGI (7 pages).
    """
    pdf     = pdfplumber.open(pdf_path)
    n_pages = len(pdf.pages)
    info    = _info(pdf)

    # ── Détection format ─────────────────────────────────
    is_dgi = (n_pages == 7)
    fmt    = "DGI" if is_dgi else "AMMC"
    logger.info(f"Format : {fmt} ({n_pages} pages)")

    if is_dgi:
        actif_pages  = [1, 2]   # pages 2-3
        passif_pages = [3]       # page 4
        cpc_pages    = [4, 5, 6] # pages 5-7
    else:
        actif_pages  = [1]       # page 2
        passif_pages = [2]       # page 3
        cpc_pages    = [3, 4]    # pages 4-5

    # ── Extraction ───────────────────────────────────────
    logger.info("Extraction Actif...")
    actif  = _get_rows(pdf, actif_pages)
    logger.info("Extraction Passif...")
    passif = _get_rows(pdf, passif_pages)
    logger.info("Extraction CPC...")
    cpc    = _get_rows(pdf, cpc_pages)
    pdf.close()

    logger.info(f"Lignes : actif={len(actif)} passif={len(passif)} cpc={len(cpc)}")

    # ── Construction Excel ────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _sheet_ident(wb, info)
    n_a = _sheet_actif(wb, info, actif)
    n_p = _sheet_passif(wb, info, passif)
    n_c = _sheet_cpc(wb, info, cpc)

    wb.save(output_path)
    total = n_a + n_p + n_c
    logger.info(f"Excel sauvegardé : {total} lignes ({n_a}a/{n_p}p/{n_c}c)")

    return {
        'tables':      3,
        'rows':        total,
        'pages':       n_pages,
        'format':      fmt,
        'info':        info,
        'exercice_fin': info.get('exercice_fin', ''),
    }
