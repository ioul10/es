"""
core/synonyms.py
Dictionnaire centralisé des synonymes de labels PDF → clé interne du template.

Structure :
    SYNONYMS = {
        'cle_interne': ['variante1 normalisée', 'variante2 normalisée', ...],
    }

La clé interne correspond exactement aux clés dans ACTIF/PASSIF/CPC des parsers.
La recherche se fait TOUJOURS dans le contexte d'un template donné :
→ si la clé interne n'existe pas dans ce template → pas de match (pas de confusion).

Pour ajouter une nouvelle variante :
    1. Identifier la clé interne du poste (voir ammc_parser.py ACTIF/PASSIF/CPC)
    2. Ajouter la variante normalisée dans la liste correspondante
    3. La normalisation = minuscules + sans accents + sans ponctuation + espaces simples
"""

import re
import unicodedata


def _n(s: str) -> str:
    """Normalisation partagée avec les parsers."""
    s = unicodedata.normalize('NFD', str(s))
    s = s.encode('ascii', 'ignore').decode().lower()
    # Retirer préfixes étoiles/points
    s = re.sub(r'^\s*[\*\.]+\s*', '', s)
    # Retirer chiffres romains isolés (sans i et v seuls car trop ambigus)
    s = re.sub(r'\b(xvi|xiv|xiii|xii|xi|ix|viii|vii|vi|iv|iii|ii)\b\s*[=\-\+\s]*', '', s, flags=re.I)
    s = re.sub(r'[^\w\s]', ' ', s)
    # Normaliser pluriels
    s = re.sub(r'\bchiffres?\b', 'chiffre', s)
    s = re.sub(r'\breserves?\b', 'reserve', s)
    s = re.sub(r'\breprises?\b', 'reprise', s)
    s = re.sub(r'\bprovisions?\b', 'provision', s)
    s = re.sub(r'\bsubventions?\b', 'subvention', s)
    s = re.sub(r'\bimpots?\b', 'impot', s)
    s = re.sub(r'\borganismes?\b', 'organisme', s)
    s = re.sub(r'\bcredits?\b', 'credit', s)
    return re.sub(r'\s+', ' ', s).strip()


# ══════════════════════════════════════════════════════════════════════════════
# DICTIONNAIRE DES SYNONYMES
# ══════════════════════════════════════════════════════════════════════════════
# Toutes les variantes sont déjà normalisées (résultat de _n())

SYNONYMS = {

    # ── IDENTIFICATION ────────────────────────────────────────────────────────
    'identifiant_fiscal': [
        'identifiant fiscal',
        'identification t v a',
        'identification tva',
        'article i s',
        'identifiant fiscal if',
        'num identifiant fiscal',
    ],
    'exercice': [
        'exercice du',
        'exercice',
        'periode du',
        'du au',
    ],

    # ── ACTIF — Immobilisé ────────────────────────────────────────────────────
    'immobilisations non valeurs': [
        'immobilisations en non valeurs',
        'immobilisation en non valeurs',
        'immobilisations non valeurs',
        'immo en non valeurs',
    ],
    'frais preliminaires': [
        'frais preliminaires',
        'frais priliminaires',
    ],
    'charges repartir': [
        'charges a repartir sur plusieurs exercices',
        'chages a repartir sur plusieurs exercices',  # faute de frappe SAPST
        'charges repartir',
    ],
    'primes remboursement obligations': [
        'primes de remboursement des obligations',
        'prime de remboursement des obligations',
    ],
    'immobilisations incorporelles': [
        'immobilisations incorporelles',
        'immobilisation incorporelle',
    ],
    'immobilisations recherche': [
        'immobilisations en recherche et developpement',
        'immobilisations en recherche et dev',
        'immo en recherche et developpement',
    ],
    'brevets marques droits': [
        'brevets marques droits et valeurs similaires',
        'brevet marques droit et valeurs similaires',
        'brevets marques droits valeurs similaires',
        'brevets  marques  droits et valeurs similaires',
    ],
    'immobilisations corporelles': [
        'immobilisations corporelles',
        'immobilisation corporelle',
    ],
    'installations techniques': [
        'installations techniques materiel et outillage',
        'installations  techniques  materiel et outillage',
        'installations techniques materiel outillage',
        'installations  techniques  materiel outillage',
    ],
    'mobilier materiel bureau': [
        'mobilier mat de bureau amenagement divers',
        'mobilier materiel de bureau et amenagement divers',
        'mobilier materiel de bureau et amenagements',
        'mobilier materiel bureau et amenagement div',
    ],
    'autres immobilisations corporelles': [
        'autres immobilisations corporelles',
        'autre immobilisation corporelle',
    ],
    'immobilisations financieres': [
        'immobilisations financieres',
        'immobilisation financiere',
    ],
    'ecarts conversion actif immobilise': [
        'ecarts de conversion actif e',
        'ecart de conversion actif',
        'ecarts de conversion actif',
        'ecarts de conversion  actif  e',
    ],
    'ecarts conversion actif circulant': [
        'ecarts de conversion actif i elements circulants',
        'ecarts de conversion actif i',
        'ecarts de conversion  actif  i',
        'ecart de conversion actif elements circulants',
        'elements circulants',
    ],
    'total i actif': [
        'total a b c d e',
        'total i a b c d e',
        'total  a b c d e',
        'total i  a b c d e',
    ],
    'stocks': [
        'stocks f',
        'stocks',
    ],
    'matieres fournitures consommables': [
        'matieres et fournitures consommables',
        'matiere et fournitures consommables',
        'matiere fournitures consommables',
    ],
    'creances actif circulant': [
        'creances de l actif circulant g',
        'creances de l actif circulant',
        'creances actif circulant',
    ],
    'fournisseurs debiteurs avances': [
        'fournis debiteurs avances et acomptes',
        'fournisseurs debiteurs avances et acomptes',
        'founisseurs debiteurs avances et accomptes',  # SAPST typo
        'fournis  debiteurs avances et acomptes',
    ],
    'personnel actif circulant': [
        'personnel debiteur',
        'personnel  debiteur',
    ],
    'etat actif': [
        'etat',
        'etat debiteur',
        'etat  debiteur',
    ],
    'comptes associes actif': [
        'comptes d associes',
        'comptes d associes debiteur',
    ],
    'comptes regularisation actif': [
        'comptes de regularisation actif',
        'comptes de regularisations  actif',
        'comptes de regularisation  actif',
    ],
    'titres valeurs placement': [
        'titres et valeurs de placement h',
        'titres valeurs de placement h',
        'titres valeurs placement h',
        'titres et valeurs de placement',
        'titres valeurs de placement',
    ],
    'total ii actif': [
        'total ii f g h i',
        'total ii  f g h i',
        'total ii f g h i ',  # SAPST avec + final
        'total f g h i',
    ],
    'banques tg ccp': [
        'banques t g et c c p',
        'banque t g et c c p',
        'banques  t g et c p debiteurs',
        'banque  t g  et c c p',
    ],
    'caisse regie avances': [
        'caisse regie d avances et accreditifs',
        'caisse regies d avances et accreditifs',
        'caisses regie d avances et accreditifs',
        'caisses  regie d avances et accreditifs',
    ],
    'total general actif': [
        'total general i ii iii',
        'total general i  ii  iii',
        'total i ii iii',
        'total l ii iii',
        'total  l  ii  iii',
    ],

    # ── PASSIF — Financement permanent ───────────────────────────────────────
    'capital social personnel': [
        'capital social ou personnel 1',
        'capital social ou personnel',
    ],
    'moins actionnaires capital': [
        'moins  actionnaires capital souscrit non appele',
        'actionnaires capital souscrit non appele',
        'moins  actionnaires  capital souscrit non appele',
    ],
    'capital appele': [
        'capital appele',
        'moins  capital appele',
    ],
    'dont verse': [
        'dont verse',
        'moins  dont verse',
    ],
    'prime emission fusion': [
        'prime d emission de fusion d apport',
        'primes d emission de fusion et d apport',
        'prime d emission  de fusion  d apport',
    ],
    'ecarts reevaluation': [
        'ecarts de reevaluation',
        'ecart de reevaluation',
        'ecart de reeevaluation',
        'ecart reevaluation',
    ],
    'reserve legale': [
        'reserve legale',
        'reserves legales',
    ],
    'report nouveau': [
        'report a nouveau 2',
        'reports a nouveau 2',
    ],
    'resultat instance affectation': [
        'resultat en instance d affectation 2',
        'resultats nets en instance d affectation 2',
        'resultat nets en instance d affectation 2',
    ],
    'resultat net exercice': [
        'resultat net de l exercice 2',
        'resultat net exercice 2',
    ],
    'subvention investissement': [
        'subvention d investissement',
        'subventions d investissement',
        'subventions d invertissement',  # faute SGTM
    ],
    'provisions reglementees': [
        'provisions reglementees',
        'provision reglementaire',
        'provisions reglementaires',
    ],
    'autres dettes financement': [
        'autres dettes de financement',
        'autre dettes de financement',  # SGTM
    ],
    'ecarts conversion passif financement': [
        'ecarts de conversion passif e',
        'ecarts de conversion  passif e',
        'ecart de conversion passif',
        'ecarts de conversion passif',
        'ecarts de conversion  passif',
    ],
    'ecarts conversion passif circulant': [
        'ecarts de conversion passif elements circulants h',
        'ecarts de conversion passif h',
        'ecarts de conversion  passif  elements circulants h',
        'ecart de conversion passif elements circulants',
        'ecarts de conversion  passif',  # 2ème occurrence
    ],
    'total i passif': [
        'total i a b c d e',
        'total i  a b c d e',
        'total  i  a b c d e',
        'total i  a b c d e ',
    ],
    'fournisseurs comptes rattaches': [
        'fournisseurs et comptes rattaches',
        'fournisseur et comptes rattaches',
    ],
    'personnel passif': [
        'personnel crediteur',
        'personnel  crediteur',
    ],
    'organismes sociaux': [
        'organismes sociaux',
        'organisme sociaux',  # BORJ/EtatsFiscaux
        'organisme social',
    ],
    'etat passif': [
        'etat crediteur',
        'etat  crediteur',
    ],
    'comptes associes passif': [
        'comptes d associes crediteurs',
        'comptes d associes  crediteur',
    ],
    'comptes regularisation passif': [
        'comptes de regularisation passif',
        'comptes de regularisations passif',
        'comptes de regularisation  passif',
    ],
    'autres provisions risques charges': [
        'autres provisions pour risques et charges g',
        'autres provision pour risques et charges',
        'autres provisions pour risques et charges',
    ],
    'total ii passif': [
        'total ii f g h',
        'total ii  f g h',
        'total  ii  f g h',
        'total f g h',
    ],
    'credits escompte': [
        'credits d escompte',
        'credit d escompte',
        'credits d escomptes',   # BORJ avec s
        'credit d escomptes',
    ],
    'credits tresorerie': [
        'credits de tresorerie',
        'credit de tresorerie',
        'credit tresorerie',
    ],
    'banques soldes crediteurs': [
        'banques soldes crediteurs',
        'banque  soldes crediteurs',
        'banques  soldes crediteurs',
        'banques  soldes  crediteurs',
    ],
    'total general passif': [
        'total general i ii iii',
        'total general i  ii  iii',
        'total i ii iii',
        'total l ii iii',
        'total i ii iii',
    ],

    # ── CPC ───────────────────────────────────────────────────────────────────
    'ventes marchandises': [
        'ventes de marchandises en l etat',
        'vente de marchandises en l etat',
        'ventes de marchandises',
        'vente de marchandises',
    ],
    'ventes biens services': [
        'ventes de biens et services produits',
        'vente de biens et services produits',
    ],
    'chiffres affaires': [
        'chiffre d affaires',
        'chiffres d affaires',
        'chiffre d affaire',
    ],
    'variation stocks produits': [
        'variation de stocks de produits 1',
        'variation des stocks de produits 1',
        'variation de stock de produits 1',
        'variation de stocks de produits',
    ],
    'immobilisations produites': [
        'immobilisations produites par l entreprise pour elle meme',
        'immobilisations produites par l entreprise',
        'immo produites par l e se por elle meme',
        'immobilisations produites par l entreprise elle meme',
    ],
    'subventions exploitation': [
        'subvention d exploitation',
        'subvention d exploitation',
        'subvention d eexploitation',   # SAPST typo
    ],
    'autres produits exploitation': [
        'autres produits d exploitation',
        'autre produit d exploitation',
    ],
    'reprises exploitation': [
        'reprise d exploitation transferts de charges',
        'reprises d exploitation  transferts de charges',
        'reprise d exploitation  transfert de charges',
        'reprises d exploitations  transfert de charges',   # SAPST
        'reprise d exploitation  transferts de charges',
        'reprise  reprise d exploitation  transferts de charges',  # BORJ double *
    ],
    'achats revendus marchandises': [
        'achats revendus 2 de marchandises',
        'achats revendus de marchandises 2',
        'achats revendus de marchandises',
        'achat revendu de marchandises 2',
    ],
    'achats consommes matieres': [
        'achats consommes 2 de matieres et fournitures',
        'achats consommes de matieres et fournitures',
        'achats consommes de matieres  de fournitures',
        'achat consommes 2 de matiere et de fournitures',  # SAPST
        'achats cosommes 2 de matiere et de fournitures',  # faute
    ],
    'autres charges externes': [
        'autres charges externes',
        'autre charges externe',
    ],
    'charges personnel': [
        'charges de personnel',
        'charge de personnel',
    ],
    'autres charges exploitation': [
        'autres charges d exploitation',
        'autre charges d exploitation',
        'autres charges d exploitaion',  # SAPST typo
    ],
    'dotations exploitation': [
        'dotations d exploitation',
        'dotation d exploitation',
    ],
    'produits titres participation': [
        'produits des titres de partic et autres titres immobilises',
        'produits des titres de participation et autres titres immobilises',
        'produits des titres de participation et des autres',
        'produits des titres de partic  et autres titres immobilises',
    ],
    'reprises financieres': [
        'reprises financieres  transferts de charges',
        'reprise financieres  transferts de charges',
        'reprises financier  transfert charges',
        'reprise financieres  transferts de changes',
        'reprise financieres  transferts de',   # tronqué SAPST
    ],
    'charges interets': [
        'charges d interets',
        'charge d interets',
        'charges d interet',
    ],
    'pertes change': [
        'pertes de change',
        'perte de change',
        'pertes change',
    ],
    'autres charges financieres': [
        'autres charges financieres',
        'autre charges financiere',
    ],
    'dotations financieres': [
        'dotations financieres',
        'dotation financiere',
    ],
    'produits cessions immobilisations': [
        'produits des cessions d immobilisations',
        'produits des cessions d immobilisation',
        'produit des cessions d immobilisation',
        'produits des cession d immobilisation',
    ],
    'reprises non courantes': [
        'reprises non courantes  transferts de charges',
        'reprise non courants  transferts de charges',
        'reprises non courantes  transferts de charges',
        'reprises non courantes  transfert de charges',
    ],
    'valeurs nettes amortissements': [
        'valeurs nettes d amortissement des immob cedees',
        'valeurs nettes d amortissements des immobilisations cedees',
        'valeurs nettes d amortissements des immob cedees',
        'valeurs nettes d amortissements des immobilisations',
        'valeur nette d amortissements des immobilisations cedees',
    ],
    'autres charges non courantes': [
        'autres charges non courantes',
        'autre charges non courante',
        'autres charges non courants',
        'autres charges non courantes',
    ],
    'dotations non courantes': [
        'dotations non courantes aux amortissements et aux provisions',
        'dotations non courantes aux amortissements',
        'dotation non courantes aux amortissements',
        'dotations non courantes',
        'dotation non courante',
    ],
    'impots resultats': [
        'impot sur les resultats',
        'impot sur les benefices',
        'mpot sur les resultats',   # après suppression du i romain
        'mpot sur les benefices',
    ],
    'resultat net': [
        'resultat net xi xii',
        'resultat net',
        'resultat net xiii xi xii',   # XIII dans le milieu du label
        'resultat net xiii  xi  xii',
    ],
    'total produits': [
        'total des produits i iv viii',
        'total des produits',
        'total des produits xiv i iv viii',   # XIV reste après regex
    ],
    'total charges': [
        'total des charges ii v ix xii',
        'total des charges',
        'xv total des charges xv ii v ix xii',  # XV reste
        'total des charges xv ii v ix xii',
    ],
    'resultat net total': [
        'resultat net total des produits  total des charges',
        'resultat net  total des produits  total des charges',
        'resultat net xvi xiv  xv',   # XVI reste
        'resultat net xiv  xv',
    ],
}

# ── Index inversé : variante → liste de clés candidates ─────────────────────
# Une même variante peut pointer vers plusieurs clés (ex: 'total l ii iii' → actif ET passif)
# lookup_in_template choisit la bonne selon le template courant
_INDEX: dict[str, list[str]] = {}
for _key, _variants in SYNONYMS.items():
    for _v in _variants:
        _v_norm = _n(_v)
        if _v_norm:
            if _v_norm not in _INDEX:
                _INDEX[_v_norm] = []
            if _key not in _INDEX[_v_norm]:
                _INDEX[_v_norm].append(_key)


def lookup(label: str) -> list[str]:
    """Retourne la liste des clés internes candidates pour un label PDF."""
    return _INDEX.get(_n(label), [])


def lookup_in_template(label: str, template: list) -> int:
    """
    Cherche l'index dans le template correspondant au label PDF via le dictionnaire.
    Essaie chaque clé candidate dans l'ordre jusqu'à en trouver une dans le template.
    Retourne l'index ou -1.
    """
    candidates = lookup(label)
    for key in candidates:
        for i, (k, _, _) in enumerate(template):
            if k == key:
                return i
    return -1
