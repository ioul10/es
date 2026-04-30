"""
core/rapport_parser.py
Parser pour rapports financiers libres (non-DGI, non-AMMC).
L'utilisateur indique manuellement les pages contenant Actif, Passif, CPC.
Réutilise le même template MCN et le même matching que dgi_parser.
"""
import re, unicodedata
from collections import defaultdict
import pdfplumber

# ── Réutilisation directe du template et matching DGI ────────────
from core.dgi_parser import (
    ACTIF, PASSIF, CPC,
    _parse, _norm, _has_fused, _detect_val_cols, _is_cpc_6col,
    _xy_rows, match_label, _build_value_map,
    _is_index_row, _is_rotated,
)
_ROMAIN_RE = re.compile(r'^(I{1,3}|IV|V?I{0,3}|IX|X{1,3}|XIV|XV|XVI)$', re.I)

# ══════════════════════════════════════════════════════════════════
# PARSING DES PAGES
# ══════════════════════════════════════════════════════════════════

def _parse_pages_input(raw: str) -> list[int]:
    """
    Convertit une saisie utilisateur en liste d'indices 0-based.
    Ex: "1"   → [0]
        "1,2" → [0, 1]
        "2-4" → [1, 2, 3]
    """
    indices = []
    raw = raw.strip()
    for part in re.split(r'[,;]', raw):
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


def _extract_section(pdf, page_indices: list, is_actif: bool = False) -> list[tuple[str, list]]:
    """
    Extrait les paires (label, [valeurs]) depuis les pages indiquées.
    Même logique que dgi_parser._extract_section avec support rapport libre.
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
                # ── CPC 6 colonnes style DGI (Nature | Label | ...) ──
                if _is_cpc_6col(t):
                    for row in t[1:]:
                        if not row: continue
                        col0  = str(row[0] or '').strip().replace('\n', ' ')
                        col1  = str(row[1] or '').strip().replace('\n', ' ')
                        label = col1 if col1 else col0
                        label = re.sub(r'^\*\s*', '', label).strip()
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

                # ── Détection rapport libre : première colonne = label ──
                val_cols = _detect_val_cols(t)
                for row in t:
                    if not row or _is_index_row(row): continue
                    if len(row) < 2: continue

                    # Essayer col0 puis col1 comme label
                    label = ''
                    for ci in [0, 1]:
                        if ci < len(row) and row[ci]:
                            candidate = str(row[ci]).strip().replace('\n', ' ')
                            candidate = re.sub(r'^\*\s*', '', candidate).strip()
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
                        if gap > 1:
                            vals = [gv(val_cols[0]), None, gv(val_cols[1]), gv(val_cols[2])]
                        else:
                            vals = [gv(val_cols[0]), gv(val_cols[1]), gv(val_cols[2]), None]
                    else:
                        vals = [gv(val_cols[i]) for i in range(min(nv, 4))]

                    if any(v is not None for v in vals):
                        rows.append((label, vals))

        all_rows.extend(rows)
    return all_rows


# ══════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════

def parse(
    pdf_path: str,
    pages_actif: str,
    pages_passif: str,
    pages_cpc: str,
    info: dict,
) -> dict:
    """
    Parse un rapport financier libre.
    
    Args:
        pdf_path   : chemin vers le PDF
        pages_actif  : ex "1" ou "1,2"
        pages_passif : ex "1"
        pages_cpc    : ex "2" ou "2,3"
        info         : dict d'identification saisi manuellement
    
    Returns:
        dict compatible avec excel_writer.write()
    """
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)

    idx_actif  = _parse_pages_input(pages_actif)
    idx_passif = _parse_pages_input(pages_passif)
    idx_cpc    = _parse_pages_input(pages_cpc)

    actif_rows  = _extract_section(pdf, idx_actif,  is_actif=True)
    passif_rows = _extract_section(pdf, idx_passif, is_actif=False)
    cpc_rows    = _extract_section(pdf, idx_cpc,    is_actif=False)
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
            'actif':  len(actif_map),
            'passif': len(passif_map),
            'cpc':    len(cpc_map),
            'actif_max':  len(ACTIF),
            'passif_max': len(PASSIF),
            'cpc_max':    len(CPC),
        }
    }
