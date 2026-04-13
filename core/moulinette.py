"""
core/moulinette.py
Remplit la moulinette standard à partir de 2 fichiers Excel FiscalXL.

Logique :
  Excel N   → valeurs ExN   → moulinette colonne F (N)
  Excel N-1 → valeurs ExN   → moulinette colonne E (N-1)
  Excel N-1 → valeurs ExN-1 → moulinette colonne D (N-2)

Colonnes FiscalXL :
  Actif  : col4 = NetN,  col5 = NetN-1
  Passif : col2 = ExN,   col3 = ExN-1
  CPC    : col4 = TotN,  col5 = TotN-1
"""
import openpyxl, re, unicodedata, shutil
from pathlib import Path

MOL_TEMPLATE = Path(__file__).parent.parent / "mol_v0.xlsx"

# ── MAPPING (row_moulinette, keyword, idx_N, idx_N1) ─────────────────────────
# Actif  vals=[Brut,Amort,NetN,NetN-1] → idx_N=2, idx_N1=3
# Passif vals=[ExN,ExN-1]              → idx_N=0, idx_N1=1
# CPC    vals=[PropN,ExPrec,TotN,TotN-1] → idx_N=2, idx_N1=3

ACTIF_MAP = [
    (9,  'frais preliminaires',                  2,3),(10,'charges repartir',2,3),
    (11, 'primes remboursement',                 2,3),(13,'immobilisations recherche',2,3),
    (14, 'brevets marques droits',               2,3),(15,'fonds commercial',2,3),
    (16, 'autres immobilisations incorporelles', 2,3),(18,'terrains',2,3),
    (19, 'constructions',                        2,3),(20,'installations techniques',2,3),
    (21, 'materiel transport',                   2,3),(22,'mobilier materiel bureau',2,3),
    (23, 'autres immobilisations corporelles',   2,3),(24,'immobilisations corporelles cours',2,3),
    (26, 'prets immobilises',                    2,3),(27,'autres creances financieres',2,3),
    (28, 'titres participation',                 2,3),(29,'autres titres immobilises',2,3),
    (31, 'diminution creances immobilisees',     2,3),(32,'augmentations dettes financement',2,3),
    (35, 'marchandises',                         2,3),(36,'matieres fournitures consommables',2,3),
    (37, 'produits cours',                       2,3),(38,'produits intermediaires residuels',2,3),
    (39, 'produits finis',                       2,3),(41,'fournisseurs debiteurs avances',2,3),
    (42, 'clients comptes rattaches',            2,3),(43,'personnel actif circulant',2,3),
    (44, 'etat actif',                           2,3),(45,'comptes associes actif',2,3),
    (46, 'autres debiteurs',                     2,3),(47,'comptes regularisation actif',2,3),
    (48, 'titres valeurs placement',             2,3),(49,'ecarts conversion actif circulant',2,3),
    (52, 'cheques valeurs encaisser',            2,3),(53,'banques tg ccp',2,3),
    (54, 'caisse regie avances',                 2,3),
]

PASSIF_MAP = [
    (10,'capital social personnel',0,1),(11,'prime emission fusion',0,1),
    (12,'ecarts reevaluation',0,1),(13,'reserve legale',0,1),
    (14,'autres reserves',0,1),(15,'report nouveau',0,1),
    (16,'resultat instance affectation',0,1),(17,'resultat net exercice',0,1),
    (19,'subvention investissement',0,1),(20,'provisions reglementees',0,1),
    (22,'emprunts obligataires',0,1),(23,'autres dettes financement',0,1),
    (26,'provisions risques',0,1),(27,'provisions charges',0,1),
    (30,'fournisseurs comptes rattaches',0,1),(31,'clients crediteurs avances',0,1),
    (32,'personnel passif',0,1),(33,'organismes sociaux',0,1),
    (34,'etat passif',0,1),(35,'comptes associes passif',0,1),
    (36,'autres creanciers',0,1),(37,'comptes regularisation passif',0,1),
    (38,'autres provisions risques charges',0,1),(39,'ecarts conversion passif',0,1),
    (42,'credits escompte',0,1),(43,'credits tresorerie',0,1),
    (44,'banques soldes crediteurs',0,1),
]

CPC_MAP = [
    (10,'ventes marchandises',2,3),(11,'ventes biens services',2,3),
    (13,'variation stocks produits',2,3),(14,'immobilisations produites',2,3),
    (15,'subventions exploitation',2,3),(16,'autres produits exploitation',2,3),
    (17,'reprises exploitation',2,3),(20,'achats revendus marchandises',2,3),
    (21,'achats consommes matieres',2,3),(22,'autres charges externes',2,3),
    (23,'impots taxes',2,3),(24,'charges personnel',2,3),
    (25,'autres charges exploitation',2,3),(26,'dotations exploitation',2,3),
    (30,'produits titres participation',2,3),(31,'gains change',2,3),
    (32,'interets autres produits financiers',2,3),(33,'reprises financieres',2,3),
    (36,'charges interets',2,3),(37,'pertes change',2,3),
    (38,'autres charges financieres',2,3),(39,'dotations financieres',2,3),
    (44,'produits cessions immobilisations',2,3),(45,'subventions equilibre',2,3),
    (46,'reprises subventions investissement',2,3),(47,'autres produits non courants',2,3),
    (48,'reprises non courantes',2,3),(51,'valeurs nettes amortissements',2,3),
    (53,'autres charges non courantes',2,3),(54,'dotations non courantes',2,3),
    (58,'impots resultats',2,3),
]


def _norm(s):
    s = unicodedata.normalize('NFD', str(s or ''))
    s = s.encode('ascii','ignore').decode().lower()
    s = re.sub(r'[^\w\s]',' ',s)
    return re.sub(r'\s+',' ',s).strip()

def _get_year(s):
    m = re.findall(r'\d{4}', str(s or ''))
    return int(m[-1]) if m else 0

def _find(data, keyword, col_idx):
    kw = set(w for w in _norm(keyword).split() if len(w)>2)
    best, val = 0, None
    for lbl, vals in data.items():
        lw = set(w for w in lbl.split() if len(w)>2)
        if not lw: continue
        s = len(kw&lw)/len(kw|lw) if kw|lw else 0
        if s > best and s >= 0.35:
            best = s
            val  = vals[col_idx] if col_idx < len(vals) else None
    return val

def read_fiscalxl(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    exercice = ''; societe = ''
    if '1 - Identification' in wb.sheetnames:
        wi = wb['1 - Identification']
        for r in range(1, wi.max_row+1):
            lbl = str(wi.cell(r,1).value or '').lower()
            val = str(wi.cell(r,2).value or '')
            if 'exercice' in lbl and not exercice: exercice = val
            if ('soci' in lbl or 'raison' in lbl) and not societe: societe = val

    def _sheet(name, cols):
        if name not in wb.sheetnames: return {}
        ws = wb[name]; data = {}
        for r in range(6, ws.max_row+1):
            lbl = ws.cell(r,1).value
            if not lbl: continue
            n = _norm(lbl)
            if len(n) < 3: continue
            data[n] = [ws.cell(r,c).value for c in cols]
        return data

    return {
        'exercice': exercice, 'societe': societe, 'annee': _get_year(exercice),
        'actif':  _sheet('2 - Bilan Actif',  [2,3,4,5]),
        'passif': _sheet('3 - Bilan Passif', [2,3]),
        'cpc':    _sheet('4 - CPC',          [2,3,4,5]),
    }

def fill_moulinette(path_n, path_n1, output_path):
    """
    path_n  : Excel année N  (la plus récente) → col F
    path_n1 : Excel année N-1                  → col E (ExN) + col D (ExN-1)
    """
    dn  = read_fiscalxl(path_n)
    dn1 = read_fiscalxl(path_n1)

    shutil.copy(str(MOL_TEMPLATE), output_path)
    wb = openpyxl.load_workbook(output_path)

    COL_F, COL_E, COL_D = 6, 5, 4
    filled = 0

    def inject(ws, row_map, section, data, use_n1):
        nonlocal filled
        for row_mol, keyword, idx_n, idx_n1 in row_map:
            v = _find(data[section], keyword, idx_n1 if use_n1 else idx_n)
            if v is not None:
                ws.cell(row_mol, COL_F if not use_n1 and data is dn
                        else COL_E if not use_n1 else COL_D).value = v
                filled += 1

    ws1,ws2,ws3 = wb['Feuil1'], wb['Feuil2'], wb['Feuil3']

    # Col F — Excel N, valeurs ExN
    for ws, mp, sec in [(ws1,ACTIF_MAP,'actif'),(ws2,PASSIF_MAP,'passif'),(ws3,CPC_MAP,'cpc')]:
        for row_mol, kw, idx_n, _ in mp:
            v = _find(dn[sec], kw, idx_n)
            if v is not None:
                ws.cell(row_mol, COL_F).value = v; filled += 1

    # Col E — Excel N-1, valeurs ExN
    for ws, mp, sec in [(ws1,ACTIF_MAP,'actif'),(ws2,PASSIF_MAP,'passif'),(ws3,CPC_MAP,'cpc')]:
        for row_mol, kw, idx_n, _ in mp:
            v = _find(dn1[sec], kw, idx_n)
            if v is not None:
                ws.cell(row_mol, COL_E).value = v; filled += 1

    # Col D — Excel N-1, valeurs ExN-1 (= N-2)
    for ws, mp, sec in [(ws1,ACTIF_MAP,'actif'),(ws2,PASSIF_MAP,'passif'),(ws3,CPC_MAP,'cpc')]:
        for row_mol, kw, _, idx_n1 in mp:
            v = _find(dn1[sec], kw, idx_n1)
            if v is not None:
                ws.cell(row_mol, COL_D).value = v; filled += 1

    # Dates en-têtes
    an = dn['annee']; an1 = dn1['annee']
    for ws, row in [(ws1,7),(ws2,8)]:
        ws.cell(row, COL_F).value = f"31/12/{an}\n(MAD)"
        ws.cell(row, COL_E).value = f"31/12/{an1}\n(MAD)"
        ws.cell(row, COL_D).value = f"31/12/{an1-1}\n(MAD)"

    wb.save(output_path)
    return {
        'filled': filled, 'societe': dn['societe'] or dn1['societe'],
        'annee_n': an, 'annee_n1': an1, 'annee_n2': an1-1,
        'exercice_n': dn['exercice'], 'exercice_n1': dn1['exercice'],
    }
