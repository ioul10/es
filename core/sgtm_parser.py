"""
sgtm_parser.py — Parser pour la liasse fiscale format SGTM (Modèle Normal IS)
Structure : 5 pages
  Page 1 : Identification (Page de garde)
  Page 2 : Bilan Actif       — Tableau N°01 (1/2)
  Page 3 : Bilan Passif      — Tableau N°01 (2/2)
  Page 4 : CPC (1/2)         — Tableau N°02 (1/2)
  Page 5 : CPC (2/2)         — Tableau N°02 (2/2)

Différence vs AMMC :
  - Page de garde avec identification complète
  - Actif sur 1 page (vs 2 pour AMMC/DGI)
  - Passif sur 1 page
  - CPC sur 2 pages
  - Les colonnes Actif : [section, libelle, Brut, Amort, Net_N, Net_N-1]
  - Les colonnes Passif : [section, libelle, N, N-1]
  - Les colonnes CPC : [section, roman, libelle, op1, op2, total_N, total_N-1]
"""

import re
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _to_float(s):
    """Convertit une chaîne au format marocain (1 234 567,89) en float."""
    if s is None:
        return 0.0
    s = str(s).strip()
    # Supprimer les espaces insécables et normaux utilisés comme séparateurs de milliers
    s = s.replace('\xa0', '').replace(' ', '').replace('\u202f', '')
    # Remplacer la virgule décimale par un point
    s = s.replace(',', '.')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _parse_cell_nums(cell):
    """Extrait tous les floats d'une cellule multi-lignes."""
    if not cell:
        return []
    nums = []
    for line in str(cell).split('\n'):
        line = line.strip()
        if line:
            v = _to_float(line)
            # On garde uniquement si la conversion a réussi (non-zéro ou était "0,00")
            if v != 0.0 or re.match(r'^[0\s,\.]+$', line.replace(' ', '').replace('\xa0', '')):
                nums.append(v)
    return nums


def _safe_get(lst, idx, default=0.0):
    try:
        return lst[idx]
    except IndexError:
        return default


def _extract_info_page1(page) -> dict:
    """Extraire les informations d'identification depuis la page de garde."""
    text = page.extract_text() or ''
    info = {}

    # Raison sociale
    m = re.search(r'Raison sociale\s*:\s*(.+)', text)
    if m:
        info['raison_sociale'] = m.group(1).strip()

    # Taxe Professionnelle
    m = re.search(r'Taxe Professionnelle\s*:\s*(\d+)', text)
    if m:
        info['taxe_pro'] = m.group(1).strip()

    # Identifiant Fiscal
    m = re.search(r'Identifiant Fiscal\s*:\s*(\d+)', text)
    if m:
        info['identifiant_fiscal'] = m.group(1).strip()

    # Adresse
    m = re.search(r'Adresse\s*:\s*(.+)', text)
    if m:
        info['adresse'] = m.group(1).strip()

    # Date de déclaration (format: A CASABLANCA Le DD/MM/YYYY)
    m = re.search(r'Le\s+(\d{2}/\d{2}/\d{4})', text)
    if m:
        info['date_declaration'] = m.group(1).strip()

    # Exercice fin — chercher dans les pages suivantes si nécessaire
    # On le complétera depuis les tableaux
    info['exercice_fin'] = '31/12/2024'  # sera mis à jour depuis entête tableaux

    return info


def _extract_exercice_from_header(page) -> str:
    """Extrait la date de fin d'exercice depuis l'entête d'un tableau."""
    text = page.extract_text() or ''
    m = re.search(r'Exercice du (\d{2}/\d{2}/\d{4}) au (\d{2}/\d{2}/\d{4})', text)
    if m:
        return m.group(2)  # date de fin
    return None


# ── Actif ──────────────────────────────────────────────────────────────────────

def _parse_actif(page) -> dict:
    """
    Parse le Bilan Actif (page 2).
    Colonnes table: [section, libelle_bloc, Brut, Amort, Net_N, Net_N-1]
    
    Mapping des lignes MCN loi 9-88 (49 lignes standard) :
    Immobilisations en non-valeurs : 3 lignes
    Immobilisations incorporelles : 4 lignes
    Immobilisations corporelles : 7 lignes
    Immobilisations financières : 4 lignes
    Écarts de conversion actif (immobilisé) : 2 lignes
    TOTAL I
    Stocks : 5 lignes
    Créances AC : 7 lignes
    Écarts de conversion AC : 1 ligne
    TOTAL II
    Trésorerie actif : 3 lignes
    TOTAL III
    TOTAL GENERAL
    """
    tables = page.extract_tables()
    if not tables:
        logger.warning("Actif SGTM : aucun tableau trouvé page 2")
        return {}

    table = tables[0]
    result = {}

    # Collecte de toutes les valeurs numériques en ordre
    # Structure: chaque row peut avoir plusieurs nums par cellule
    # On aplatit dans l'ordre des rows et dans l'ordre des nums dans chaque cellule

    def collect_col(col_idx):
        """Collecte tous les floats de la colonne col_idx dans l'ordre."""
        vals = []
        for row in table:
            if len(row) > col_idx:
                vals.extend(_parse_cell_nums(row[col_idx]))
        return vals

    brut  = collect_col(2)
    amort = collect_col(3)
    net_n = collect_col(4)
    net_1 = collect_col(5)

    logger.debug(f"Actif SGTM - brut[{len(brut)}]: {brut[:5]}...")
    logger.debug(f"Actif SGTM - net_n[{len(net_n)}]: {net_n[:5]}...")

    # ── MAPPING EXACT basé sur l'analyse du PDF SGTM ──────────────────────
    # Les valeurs arrivent dans cet ordre précis dans la colonne
    # (vérifié par analyse pdfplumber row-by-row)

    # IMMOBILISATIONS EN NON-VALEURS (rows 3-4 → 3 valeurs détail + 1 total)
    # Row3: Frais prélim BRUT=10M, Row4: [0,10M,0] = [Frais,TOTAL_NV,Prime]
    # Ordre réel: Frais préliminaires, TOTAL_NV, Prime remboursement
    # puis INCORPORELLES row5=total, row6=[0,total,0,0]=[RI&D,TOTAL_INC,Brevets,Fonds]
    # Note: les totaux de section apparaissent en double (once seul, once dans multi)
    # On prend la version "seule" pour les totaux, et les multis pour les détails

    # La stratégie la plus robuste : collecter UNIQUEMENT les valeurs "seules"
    # (rows dont la cellule ne contient qu'une seule valeur) = lignes de détail
    # et les valeurs "en paquet" pour les lignes multiples.

    # Reconstruction par pattern reconnu dans le PDF analysé :
    # brut: [10M, 0,10M,0, 6.7M, 0,6.7M,0,0, 922M, 5.3M,17.9M,644.8M,30.5M,175.8M,0,47.7M,
    #        108.6M, 0,4.2M,104.4M,0, 0, 0,0, 1047M,
    #        236.7M, 20.97M,207.1M,0,0,8.55M, 8236M,
    #        291.2M,6482.7M,27.9M,793M,531.6M,49.1M,61M,
    #        476.2M, 10.7M, 8960M,
    #        733.8M, 0,733.6M,0.2M, 733.8M,
    #        10741M]

    # On utilise un mapping positionnel par index dans la liste aplatie

    # ── Immobilisations en non-valeurs ─────────────────────────────────────
    # idx 0: Frais préliminaires (row3 seul)
    # idx 1,2,3 = row4 multi: [0, 10M, 0] = Charges à répartir, TOTAL_NV, Prime remboursement
    # Mais en réalité row4 = [Charges répartir, TOTAL_NV, Prime remboursement]
    # D'après l'analyse: row4 brut = [0.0, 10000000.0, 0.0]
    # L'ordre des libellés dans le PDF: Frais prélim / Charges répartir / Prime remboursement / TOTAL

    r = {}

    # Pour éviter les ambiguïtés, on re-parse row par row de façon ciblée
    def row_vals(row_idx, col_idx):
        if row_idx < len(table) and col_idx < len(table[row_idx]):
            return _parse_cell_nums(table[row_idx][col_idx])
        return []

    # ── ACTIF IMMOBILISE ───────────────────────────────────────────────────
    # Immobilisations en non-valeurs
    # Row 3 (seul): Frais préliminaires
    # Row 4 (multi 3): Charges à répartir / TOTAL_NV / Prime remboursement
    r['nv_frais_prelim_brut']     = _safe_get(row_vals(3,2), 0)
    r['nv_frais_prelim_net']      = _safe_get(row_vals(3,4), 0)
    r['nv_frais_prelim_net1']     = _safe_get(row_vals(3,5), 0)

    _r4b = row_vals(4,2); _r4n = row_vals(4,4); _r4n1 = row_vals(4,5)
    r['nv_charges_repartir_brut'] = _safe_get(_r4b, 0)
    r['nv_charges_repartir_net']  = _safe_get(_r4n, 0)
    r['nv_charges_repartir_net1'] = _safe_get(_r4n1, 0)
    r['nv_total_brut']            = _safe_get(_r4b, 1)
    r['nv_total_net']             = _safe_get(_r4n, 1)
    r['nv_total_net1']            = _safe_get(_r4n1, 1)
    r['nv_prime_remboursement_brut'] = _safe_get(_r4b, 2)
    r['nv_prime_remboursement_net']  = _safe_get(_r4n, 2)
    r['nv_prime_remboursement_net1'] = _safe_get(_r4n1, 2)

    # Immobilisations incorporelles
    # Row 5 (seul): TOTAL incorporelles
    # Row 6 (multi 4): RI&D / TOTAL_INC / Brevets / Fonds commercial
    r['inc_total_brut']           = _safe_get(row_vals(5,2), 0)
    r['inc_total_net']            = _safe_get(row_vals(5,4), 0)
    r['inc_total_net1']           = _safe_get(row_vals(5,5), 0)

    _r6b = row_vals(6,2); _r6n = row_vals(6,4); _r6n1 = row_vals(6,5)
    r['inc_rd_brut']              = _safe_get(_r6b, 0)
    r['inc_rd_net']               = _safe_get(_r6n, 0)
    r['inc_rd_net1']              = _safe_get(_r6n1, 0)
    r['inc_brevets_brut']         = _safe_get(_r6b, 2)
    r['inc_brevets_net']          = _safe_get(_r6n, 2)
    r['inc_brevets_net1']         = _safe_get(_r6n1, 2)
    r['inc_fonds_commercial_brut']= _safe_get(_r6b, 3)
    r['inc_fonds_commercial_net'] = _safe_get(_r6n, 3)
    r['inc_fonds_commercial_net1']= _safe_get(_r6n1, 3)

    # Immobilisations corporelles
    # Row 7 (seul): TOTAL corporelles
    # Row 8 (multi 7): Terrains/Constr/Instal/Mater.transp/Mobilier/Autres/En cours
    r['corp_total_brut']          = _safe_get(row_vals(7,2), 0)
    r['corp_total_net']           = _safe_get(row_vals(7,4), 0)
    r['corp_total_net1']          = _safe_get(row_vals(7,5), 0)

    _r8b = row_vals(8,2); _r8n = row_vals(8,4); _r8n1 = row_vals(8,5)
    r['corp_terrains_brut']       = _safe_get(_r8b, 0)
    r['corp_terrains_net']        = _safe_get(_r8n, 0)
    r['corp_terrains_net1']       = _safe_get(_r8n1, 0)
    r['corp_constructions_brut']  = _safe_get(_r8b, 1)
    r['corp_constructions_net']   = _safe_get(_r8n, 1)
    r['corp_constructions_net1']  = _safe_get(_r8n1, 1)
    r['corp_instal_brut']         = _safe_get(_r8b, 2)
    r['corp_instal_net']          = _safe_get(_r8n, 2)
    r['corp_instal_net1']         = _safe_get(_r8n1, 2)
    r['corp_mater_transport_brut']= _safe_get(_r8b, 3)
    r['corp_mater_transport_net'] = _safe_get(_r8n, 3)
    r['corp_mater_transport_net1']= _safe_get(_r8n1, 3)
    r['corp_mobilier_brut']       = _safe_get(_r8b, 4)
    r['corp_mobilier_net']        = _safe_get(_r8n, 4)
    r['corp_mobilier_net1']       = _safe_get(_r8n1, 4)
    r['corp_autres_brut']         = _safe_get(_r8b, 5)
    r['corp_autres_net']          = _safe_get(_r8n, 5)
    r['corp_autres_net1']         = _safe_get(_r8n1, 5)
    r['corp_en_cours_brut']       = _safe_get(_r8b, 6)
    r['corp_en_cours_net']        = _safe_get(_r8n, 6)
    r['corp_en_cours_net1']       = _safe_get(_r8n1, 6)

    # Immobilisations financières
    # Row 9 (seul): TOTAL financières
    # Row 10 (multi 4): Prêts / TOTAL_FIN / Titres participation / Autres titres
    r['fin_total_brut']           = _safe_get(row_vals(9,2), 0)
    r['fin_total_net']            = _safe_get(row_vals(9,4), 0)
    r['fin_total_net1']           = _safe_get(row_vals(9,5), 0)

    _r10b = row_vals(10,2); _r10n = row_vals(10,4); _r10n1 = row_vals(10,5)
    r['fin_prets_brut']           = _safe_get(_r10b, 0)
    r['fin_prets_net']            = _safe_get(_r10n, 0)
    r['fin_prets_net1']           = _safe_get(_r10n1, 0)
    r['fin_titres_participation_brut'] = _safe_get(_r10b, 2)
    r['fin_titres_participation_net']  = _safe_get(_r10n, 2)
    r['fin_titres_participation_net1'] = _safe_get(_r10n1, 2)
    r['fin_autres_titres_brut']   = _safe_get(_r10b, 3)
    r['fin_autres_titres_net']    = _safe_get(_r10n, 3)
    r['fin_autres_titres_net1']   = _safe_get(_r10n1, 3)

    # Écarts de conversion actif immobilisé
    # Row 11 (seul): TOTAL écarts
    # Row 12 (multi 2): Diminution créances / Aug dettes financement
    r['eca_immo_total_brut']      = _safe_get(row_vals(11,2), 0)
    r['eca_immo_total_net']       = _safe_get(row_vals(11,4), 0)
    r['eca_immo_total_net1']      = _safe_get(row_vals(11,5), 0)

    # TOTAL I
    r['total_i_brut']             = _safe_get(row_vals(13,2), 0)
    r['total_i_net']              = _safe_get(row_vals(13,4), 0)
    r['total_i_net1']             = _safe_get(row_vals(13,5), 0)

    # ── ACTIF CIRCULANT ────────────────────────────────────────────────────
    # Stocks
    # Row 14 (seul): TOTAL stocks
    # Row 15 (multi 5): Marchandises/Matieres/Produits cours/Produits finis/Autres
    r['stocks_total_brut']        = _safe_get(row_vals(14,2), 0)
    r['stocks_total_net']         = _safe_get(row_vals(14,4), 0)
    r['stocks_total_net1']        = _safe_get(row_vals(14,5), 0)

    _r15b = row_vals(15,2); _r15n = row_vals(15,4); _r15n1 = row_vals(15,5)
    r['stocks_marchandises_brut'] = _safe_get(_r15b, 0)
    r['stocks_marchandises_net']  = _safe_get(_r15n, 0)
    r['stocks_marchandises_net1'] = _safe_get(_r15n1, 0)
    r['stocks_matieres_brut']     = _safe_get(_r15b, 1)
    r['stocks_matieres_net']      = _safe_get(_r15n, 1)
    r['stocks_matieres_net1']     = _safe_get(_r15n1, 1)
    r['stocks_produits_finis_brut']= _safe_get(_r15b, 4)
    r['stocks_produits_finis_net'] = _safe_get(_r15n, 4)
    r['stocks_produits_finis_net1']= _safe_get(_r15n1, 4)

    # Créances actif circulant
    # Row 16 (seul): TOTAL créances AC
    # Row 17 (multi 7): Fournisseurs/Clients/Perso/État/Associés/Autres/Régul
    r['cre_ac_total_brut']        = _safe_get(row_vals(16,2), 0)
    r['cre_ac_total_net']         = _safe_get(row_vals(16,4), 0)
    r['cre_ac_total_net1']        = _safe_get(row_vals(16,5), 0)

    _r17b = row_vals(17,2); _r17n = row_vals(17,4); _r17n1 = row_vals(17,5)
    r['cre_ac_fournisseurs_brut'] = _safe_get(_r17b, 0)
    r['cre_ac_fournisseurs_net']  = _safe_get(_r17n, 0)
    r['cre_ac_fournisseurs_net1'] = _safe_get(_r17n1, 0)
    r['cre_ac_clients_brut']      = _safe_get(_r17b, 1)
    r['cre_ac_clients_net']       = _safe_get(_r17n, 1)
    r['cre_ac_clients_net1']      = _safe_get(_r17n1, 1)
    r['cre_ac_etat_brut']         = _safe_get(_r17b, 4)
    r['cre_ac_etat_net']          = _safe_get(_r17n, 4)
    r['cre_ac_etat_net1']         = _safe_get(_r17n1, 4)
    r['cre_ac_autres_brut']       = _safe_get(_r17b, 5)
    r['cre_ac_autres_net']        = _safe_get(_r17n, 5)
    r['cre_ac_autres_net1']       = _safe_get(_r17n1, 5)

    # Titres et valeurs de placement
    r['tvp_total_brut']           = _safe_get(row_vals(18,2), 0)
    r['tvp_total_net']            = _safe_get(row_vals(18,4), 0)
    r['tvp_total_net1']           = _safe_get(row_vals(18,5), 0)

    # Écarts de conversion actif circulant
    r['eca_ac_total_brut']        = _safe_get(row_vals(19,2), 0)
    r['eca_ac_total_net']         = _safe_get(row_vals(19,4), 0)
    r['eca_ac_total_net1']        = _safe_get(row_vals(19,5), 0)

    # TOTAL II
    r['total_ii_brut']            = _safe_get(row_vals(20,2), 0)
    r['total_ii_net']             = _safe_get(row_vals(20,4), 0)
    r['total_ii_net1']            = _safe_get(row_vals(20,5), 0)

    # ── TRÉSORERIE ACTIF ───────────────────────────────────────────────────
    # Row 21 (seul): TOTAL tréso
    # Row 22 (multi 3): Chèques / Banque / Caisses
    r['treso_total_brut']         = _safe_get(row_vals(21,2), 0)
    r['treso_total_net']          = _safe_get(row_vals(21,4), 0)
    r['treso_total_net1']         = _safe_get(row_vals(21,5), 0)

    _r22b = row_vals(22,2); _r22n = row_vals(22,4); _r22n1 = row_vals(22,5)
    r['treso_cheques_brut']       = _safe_get(_r22b, 0)
    r['treso_cheques_net']        = _safe_get(_r22n, 0)
    r['treso_cheques_net1']       = _safe_get(_r22n1, 0)
    r['treso_banque_brut']        = _safe_get(_r22b, 1)
    r['treso_banque_net']         = _safe_get(_r22n, 1)
    r['treso_banque_net1']        = _safe_get(_r22n1, 1)
    r['treso_caisses_brut']       = _safe_get(_r22b, 2)
    r['treso_caisses_net']        = _safe_get(_r22n, 2)
    r['treso_caisses_net1']       = _safe_get(_r22n1, 2)

    # TOTAL III = TOTAL TRÉSO (row 23 = répétition du row 21)
    r['total_iii_brut']           = _safe_get(row_vals(23,2), 0)
    r['total_iii_net']            = _safe_get(row_vals(23,4), 0)
    r['total_iii_net1']           = _safe_get(row_vals(23,5), 0)

    # TOTAL GÉNÉRAL
    r['total_general_brut']       = _safe_get(row_vals(24,2), 0)
    r['total_general_net']        = _safe_get(row_vals(24,4), 0)
    r['total_general_net1']       = _safe_get(row_vals(24,5), 0)

    return r


# ── Passif ─────────────────────────────────────────────────────────────────────

def _parse_passif(page) -> dict:
    """
    Parse le Bilan Passif (page 3).
    Colonnes: [section, libelle_bloc, N, N-1]
    """
    tables = page.extract_tables()
    if not tables:
        logger.warning("Passif SGTM : aucun tableau trouvé page 3")
        return {}

    table = tables[0]
    r = {}

    def rv(row_idx, col_idx):
        if row_idx < len(table) and col_idx < len(table[row_idx]):
            return _parse_cell_nums(table[row_idx][col_idx])
        return []

    # ── FINANCEMENT PERMANENT ──────────────────────────────────────────────
    # Row 1 (seul): TOTAL Capitaux propres
    # Row 2 (multi 11): Capital/Moins/Appelé/Versé/Prime/Réesval/Rés.lég/Autres/Reports/Résult.inst/Résult.net
    # Row 3 (répétition TOTAL CP)
    r['cp_total_n']               = _safe_get(rv(1,2), 0)
    r['cp_total_n1']              = _safe_get(rv(1,3), 0)

    _r2n = rv(2,2); _r2n1 = rv(2,3)
    r['cp_capital_n']             = _safe_get(_r2n, 0)
    r['cp_capital_n1']            = _safe_get(_r2n1, 0)
    r['cp_capital_appele_n']      = _safe_get(_r2n, 2)
    r['cp_capital_appele_n1']     = _safe_get(_r2n1, 2)
    r['cp_reserve_legale_n']      = _safe_get(_r2n, 6)
    r['cp_reserve_legale_n1']     = _safe_get(_r2n1, 6)
    r['cp_autres_reserves_n']     = _safe_get(_r2n, 7)
    r['cp_autres_reserves_n1']    = _safe_get(_r2n1, 7)
    r['cp_reports_nouveau_n']     = _safe_get(_r2n, 8)
    r['cp_reports_nouveau_n1']    = _safe_get(_r2n1, 8)
    r['cp_resultat_net_n']        = _safe_get(_r2n, 10)
    r['cp_resultat_net_n1']       = _safe_get(_r2n1, 10)

    # Capitaux propres assimilés (B)
    r['cpa_total_n']              = _safe_get(rv(4,2), 0)
    r['cpa_total_n1']             = _safe_get(rv(4,3), 0)
    _r5n = rv(5,2); _r5n1 = rv(5,3)
    r['cpa_subventions_n']        = _safe_get(_r5n, 0)
    r['cpa_subventions_n1']       = _safe_get(_r5n1, 0)
    r['cpa_prov_reglementees_n']  = _safe_get(_r5n, 1)
    r['cpa_prov_reglementees_n1'] = _safe_get(_r5n1, 1)

    # Dettes de financement (C)
    r['dfi_total_n']              = _safe_get(rv(6,2), 0)
    r['dfi_total_n1']             = _safe_get(rv(6,3), 0)
    _r7n = rv(7,2); _r7n1 = rv(7,3)
    r['dfi_emprunts_oblig_n']     = _safe_get(_r7n, 0)
    r['dfi_emprunts_oblig_n1']    = _safe_get(_r7n1, 0)
    r['dfi_autres_dettes_n']      = _safe_get(_r7n, 1)
    r['dfi_autres_dettes_n1']     = _safe_get(_r7n1, 1)

    # Provisions durables risques et charges (D)
    r['pdrc_total_n']             = _safe_get(rv(8,2), 0)
    r['pdrc_total_n1']            = _safe_get(rv(8,3), 0)
    _r9n = rv(9,2); _r9n1 = rv(9,3)
    r['pdrc_risques_n']           = _safe_get(_r9n, 0)
    r['pdrc_risques_n1']          = _safe_get(_r9n1, 0)
    r['pdrc_charges_n']           = _safe_get(_r9n, 1)
    r['pdrc_charges_n1']          = _safe_get(_r9n1, 1)

    # Écarts de conversion passif (E)
    r['ecp_total_n']              = _safe_get(rv(10,2), 0)
    r['ecp_total_n1']             = _safe_get(rv(10,3), 0)

    # TOTAL I
    r['total_i_n']                = _safe_get(rv(12,2), 0)
    r['total_i_n1']               = _safe_get(rv(12,3), 0)

    # ── PASSIF CIRCULANT ───────────────────────────────────────────────────
    # Row 13 (seul): TOTAL dettes passif circulant (F)
    # Row 14 (multi 8): Fournisseurs/Clients créditeurs/Perso/Orga.soc/État/Associés/Autres/Régul
    r['dpc_total_n']              = _safe_get(rv(13,2), 0)
    r['dpc_total_n1']             = _safe_get(rv(13,3), 0)

    _r14n = rv(14,2); _r14n1 = rv(14,3)
    r['dpc_fournisseurs_n']       = _safe_get(_r14n, 0)
    r['dpc_fournisseurs_n1']      = _safe_get(_r14n1, 0)
    r['dpc_clients_credit_n']     = _safe_get(_r14n, 1)
    r['dpc_clients_credit_n1']    = _safe_get(_r14n1, 1)
    r['dpc_personnel_n']          = _safe_get(_r14n, 2)
    r['dpc_personnel_n1']         = _safe_get(_r14n1, 2)
    r['dpc_orga_sociaux_n']       = _safe_get(_r14n, 3)
    r['dpc_orga_sociaux_n1']      = _safe_get(_r14n1, 3)
    r['dpc_etat_n']               = _safe_get(_r14n, 4)
    r['dpc_etat_n1']              = _safe_get(_r14n1, 4)
    r['dpc_comptes_assoc_n']      = _safe_get(_r14n, 5)
    r['dpc_comptes_assoc_n1']     = _safe_get(_r14n1, 5)
    r['dpc_autres_n']             = _safe_get(_r14n, 6)
    r['dpc_autres_n1']            = _safe_get(_r14n1, 6)
    r['dpc_regularisation_n']     = _safe_get(_r14n, 7)
    r['dpc_regularisation_n1']    = _safe_get(_r14n1, 7)

    # Autres provisions (G)
    r['aprc_total_n']             = _safe_get(rv(15,2), 0)
    r['aprc_total_n1']            = _safe_get(rv(15,3), 0)

    # Écarts de conversion passif circulant (H)
    r['ecpc_total_n']             = _safe_get(rv(16,2), 0)
    r['ecpc_total_n1']            = _safe_get(rv(16,3), 0)

    # TOTAL II
    r['total_ii_n']               = _safe_get(rv(17,2), 0)
    r['total_ii_n1']              = _safe_get(rv(17,3), 0)

    # ── TRÉSORERIE PASSIF ──────────────────────────────────────────────────
    r['tp_total_n']               = _safe_get(rv(18,2), 0)
    r['tp_total_n1']              = _safe_get(rv(18,3), 0)

    _r19n = rv(19,2); _r19n1 = rv(19,3)
    r['tp_credit_escompte_n']     = _safe_get(_r19n, 0)
    r['tp_credit_escompte_n1']    = _safe_get(_r19n1, 0)
    r['tp_credit_tresorerie_n']   = _safe_get(_r19n, 1)
    r['tp_credit_tresorerie_n1']  = _safe_get(_r19n1, 1)
    r['tp_banques_n']             = _safe_get(_r19n, 2)
    r['tp_banques_n1']            = _safe_get(_r19n1, 2)

    # TOTAL III (row 20 = répétition)
    r['total_iii_n']              = _safe_get(rv(20,2), 0)
    r['total_iii_n1']             = _safe_get(rv(20,3), 0)

    # TOTAL GÉNÉRAL
    r['total_general_n']          = _safe_get(rv(21,2), 0)
    r['total_general_n1']         = _safe_get(rv(21,3), 0)

    return r


# ── CPC ────────────────────────────────────────────────────────────────────────

def _parse_cpc(page4, page5) -> dict:
    """
    Parse le CPC (pages 4 et 5).
    Col structure: [section, roman, libelle, op_propre, op_precedent, total_N, total_N1]
    On utilise la colonne 5 (total_N) et colonne 6 (total_N1).
    """
    tables4 = page4.extract_tables()
    tables5 = page5.extract_tables()
    if not tables4 or not tables5:
        logger.warning("CPC SGTM : tableaux manquants")
        return {}

    t4 = tables4[0]
    t5 = tables5[0]
    r = {}

    def rv4(row_idx): 
        n  = _parse_cell_nums(t4[row_idx][5]) if row_idx < len(t4) and len(t4[row_idx]) > 5 else []
        n1 = _parse_cell_nums(t4[row_idx][6]) if row_idx < len(t4) and len(t4[row_idx]) > 6 else []
        return n, n1

    def rv5(row_idx):
        n  = _parse_cell_nums(t5[row_idx][5]) if row_idx < len(t5) and len(t5[row_idx]) > 5 else []
        n1 = _parse_cell_nums(t5[row_idx][6]) if row_idx < len(t5) and len(t5[row_idx]) > 6 else []
        return n, n1

    # ── PAGE 4 — EXPLOITATION + FINANCIER ─────────────────────────────────
    # Row 2: [0, Ventes biens] → col5=[0, CA]
    _n, _n1 = rv4(2)
    r['ventes_marchandises_n']     = _safe_get(_n,  0)
    r['ventes_marchandises_n1']    = _safe_get(_n1, 0)
    r['ventes_biens_services_n']   = _safe_get(_n,  1)
    r['ventes_biens_services_n1']  = _safe_get(_n1, 1)

    # Row 3: [CA, Variation stocks, Immo prod, Subv, Autres, Reprises] → 6 vals
    _n, _n1 = rv4(3)
    r['chiffre_affaires_n']        = _safe_get(_n,  0)
    r['chiffre_affaires_n1']       = _safe_get(_n1, 0)
    r['variation_stocks_produits_n']= _safe_get(_n, 1)
    r['variation_stocks_produits_n1']= _safe_get(_n1,1)
    r['autres_produits_exploit_n'] = _safe_get(_n,  4)
    r['autres_produits_exploit_n1']= _safe_get(_n1, 4)
    r['reprises_exploit_n']        = _safe_get(_n,  5)
    r['reprises_exploit_n1']       = _safe_get(_n1, 5)

    # Row 4: TOTAL I (produits exploitation)
    _n, _n1 = rv4(4)
    r['total_produits_exploit_n']  = _safe_get(_n,  0)
    r['total_produits_exploit_n1'] = _safe_get(_n1, 0)

    # Row 5: Charges exploitation [0, Achats consommés, Autres charges ext, Impôts,
    #                               Charges perso, Autres charges, Dot exploit] → 7 vals
    _n, _n1 = rv4(5)
    r['achats_revendus_n']         = _safe_get(_n,  0)
    r['achats_revendus_n1']        = _safe_get(_n1, 0)
    r['achats_consommes_n']        = _safe_get(_n,  1)
    r['achats_consommes_n1']       = _safe_get(_n1, 1)
    r['autres_charges_ext_n']      = _safe_get(_n,  2)
    r['autres_charges_ext_n1']     = _safe_get(_n1, 2)
    r['impots_taxes_n']            = _safe_get(_n,  3)
    r['impots_taxes_n1']           = _safe_get(_n1, 3)
    r['charges_personnel_n']       = _safe_get(_n,  4)
    r['charges_personnel_n1']      = _safe_get(_n1, 4)
    r['autres_charges_exploit_n']  = _safe_get(_n,  5)
    r['autres_charges_exploit_n1'] = _safe_get(_n1, 5)
    r['dot_exploitation_n']        = _safe_get(_n,  6)
    r['dot_exploitation_n1']       = _safe_get(_n1, 6)

    # Row 6: [TOTAL II, RÉSULTAT EXPLOITATION] → 2 vals
    _n, _n1 = rv4(6)
    r['total_charges_exploit_n']   = _safe_get(_n,  0)
    r['total_charges_exploit_n1']  = _safe_get(_n1, 0)
    r['resultat_exploitation_n']   = _safe_get(_n,  1)
    r['resultat_exploitation_n1']  = _safe_get(_n1, 1)

    # Row 7: Produits financiers [Produits titres, Gains change, Intérêts, Reprises fin] → 4 vals
    _n, _n1 = rv4(7)
    r['produits_titres_n']         = _safe_get(_n,  0)
    r['produits_titres_n1']        = _safe_get(_n1, 0)
    r['gains_change_n']            = _safe_get(_n,  1)
    r['gains_change_n1']           = _safe_get(_n1, 1)
    r['interets_produits_n']       = _safe_get(_n,  2)
    r['interets_produits_n1']      = _safe_get(_n1, 2)
    r['reprises_financieres_n']    = _safe_get(_n,  3)
    r['reprises_financieres_n1']   = _safe_get(_n1, 3)

    # Row 8: TOTAL IV (produits financiers)
    _n, _n1 = rv4(8)
    r['total_produits_fin_n']      = _safe_get(_n,  0)
    r['total_produits_fin_n1']     = _safe_get(_n1, 0)

    # Row 9: Charges financières [Intérêts charges, Pertes change, Autres ch.fin, Dot fin] → 4 vals
    _n, _n1 = rv4(9)
    r['charges_interets_n']        = _safe_get(_n,  0)
    r['charges_interets_n1']       = _safe_get(_n1, 0)
    r['pertes_change_n']           = _safe_get(_n,  1)
    r['pertes_change_n1']          = _safe_get(_n1, 1)
    r['autres_charges_fin_n']      = _safe_get(_n,  2)
    r['autres_charges_fin_n1']     = _safe_get(_n1, 2)
    r['dot_financieres_n']         = _safe_get(_n,  3)
    r['dot_financieres_n1']        = _safe_get(_n1, 3)

    # Row 10: [TOTAL V, RÉSULTAT FINANCIER]
    _n, _n1 = rv4(10)
    r['total_charges_fin_n']       = _safe_get(_n,  0)
    r['total_charges_fin_n1']      = _safe_get(_n1, 0)
    r['resultat_financier_n']      = _safe_get(_n,  1)
    r['resultat_financier_n1']     = _safe_get(_n1, 1)

    # ── PAGE 5 — NON COURANT ───────────────────────────────────────────────
    # Row 0: RÉSULTAT COURANT (en-tête de colonne)
    _n, _n1 = rv5(0)
    # Row 0 col5 contient "TOTAUX DE L'EXERCICE \n 3=1+2 \n 1130474654.45"
    # Le résultat courant est la 1ère vraie valeur numérique (après le texte d'entête)
    r['resultat_courant_n']        = _safe_get(_n,  0)
    # Pour N-1, la colonne 6 a "4 \n 775326574.13"
    # _safe_get(_n1, ...) peut prendre l'entier 4 comme premier élément
    r['resultat_courant_n1']       = _safe_get(_n1, 1) if len(_n1) > 1 else _safe_get(_n1, 0)

    # Row 2: Produits non courants [Cessions, Subv équil, Repr.subv, Autres prod NC, Repr NC] → 6 vals
    _n, _n1 = rv5(2)
    r['produits_cessions_n']       = _safe_get(_n,  0)
    r['produits_cessions_n1']      = _safe_get(_n1, 0)
    r['autres_produits_nc_n']      = _safe_get(_n,  3)
    r['autres_produits_nc_n1']     = _safe_get(_n1, 3)
    r['reprises_nc_n']             = _safe_get(_n,  5)
    r['reprises_nc_n1']            = _safe_get(_n1, 5)

    # Row 3: TOTAL VIII (produits non courants)
    _n, _n1 = rv5(3)
    r['total_produits_nc_n']       = _safe_get(_n,  0)
    r['total_produits_nc_n1']      = _safe_get(_n1, 0)

    # Row 4: Charges non courantes [VNA immo, Subv accordées, Autres, Dot NC] → 5 vals
    _n, _n1 = rv5(4)
    r['vna_immobilisations_n']     = _safe_get(_n,  0)
    r['vna_immobilisations_n1']    = _safe_get(_n1, 0)
    r['subventions_accordees_n']   = _safe_get(_n,  1)
    r['subventions_accordees_n1']  = _safe_get(_n1, 1)
    r['autres_charges_nc_n']       = _safe_get(_n,  2)
    r['autres_charges_nc_n1']      = _safe_get(_n1, 2)
    r['dot_nc_amort_prov_n']       = _safe_get(_n,  3)
    r['dot_nc_amort_prov_n1']      = _safe_get(_n1, 3)

    # Row 5: [TOTAL IX, RÉSULTAT NON COURANT]
    _n, _n1 = rv5(5)
    r['total_charges_nc_n']        = _safe_get(_n,  0)
    r['total_charges_nc_n1']       = _safe_get(_n1, 0)
    r['resultat_nc_n']             = _safe_get(_n,  1)
    r['resultat_nc_n1']            = _safe_get(_n1, 1)

    # Row 6: RÉSULTAT AVANT IMPÔTS
    _n, _n1 = rv5(6)
    r['resultat_avant_impots_n']   = _safe_get(_n,  0)
    r['resultat_avant_impots_n1']  = _safe_get(_n1, 0)

    # Row 7: [IMPÔTS, RÉSULTAT NET]
    _n, _n1 = rv5(7)
    r['impots_resultats_n']        = _safe_get(_n,  0)
    r['impots_resultats_n1']       = _safe_get(_n1, 0)
    r['resultat_net_n']            = _safe_get(_n,  1)
    r['resultat_net_n1']           = _safe_get(_n1, 1)

    # Totaux généraux (lignes XVI-XVII dans l'ordre standard)
    # Total produits = Total I + IV + VIII
    r['total_produits_n']          = (r['total_produits_exploit_n'] +
                                      r['total_produits_fin_n'] +
                                      r['total_produits_nc_n'])
    r['total_charges_n']           = (r['total_charges_exploit_n'] +
                                      r['total_charges_fin_n'] +
                                      r['total_charges_nc_n'] +
                                      r['impots_resultats_n'])

    return r


# ── Entrée publique ─────────────────────────────────────────────────────────────

def parse(pdf_path: str) -> dict:
    """
    Parse une liasse fiscale au format SGTM (5 pages).
    Retourne un dict avec les clés : info, actif, passif, cpc.
    """
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 5:
            raise ValueError(f"PDF SGTM attendu : 5 pages, trouvé {len(pdf.pages)} page(s)")

        logger.info(f"SGTM Parser : {len(pdf.pages)} pages détectées")

        # Identification
        info = _extract_info_page1(pdf.pages[0])

        # Date exercice depuis entête tableau
        exercice_fin = _extract_exercice_from_header(pdf.pages[1])
        if exercice_fin:
            info['exercice_fin'] = exercice_fin

        # Tableaux financiers
        actif  = _parse_actif(pdf.pages[1])
        passif = _parse_passif(pdf.pages[2])
        cpc    = _parse_cpc(pdf.pages[3], pdf.pages[4])

    logger.info(f"SGTM Parser OK — {info.get('raison_sociale','?')} {info.get('exercice_fin','?')}")

    return {
        'format': 'SGTM',
        'info':   info,
        'actif':  actif,
        'passif': passif,
        'cpc':    cpc,
    }
