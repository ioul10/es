"""
FiscalXL Pro — Convertisseur PDF fiscal → Excel structuré
Premier outil de la suite de notation des entreprises marocaines (MCN loi 9-88)
"""
import streamlit as st
import tempfile, os
from utils.logger import get_logger

logger = get_logger(__name__)

st.set_page_config(
    page_title="FiscalXL Pro",
    page_icon="📊",
    layout="wide"
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""<style>
.hdr{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.6rem 2rem;
border-radius:12px;margin-bottom:1.4rem;}
.hdr h1{color:white;margin:0;font-size:2rem;letter-spacing:.5px;}
.hdr p{color:#BDD7EE;margin:.4rem 0 0;font-size:.95rem;}
.welcome{background:#f8fafd;border-left:4px solid #2E75B6;
border-radius:0 10px 10px 0;padding:1.2rem 1.6rem;margin-bottom:1.2rem;}
.welcome h3{color:#1F3864;margin:0 0 .5rem;}
.welcome p{color:#444;margin:0;line-height:1.6;}
.step-badge{display:inline-block;background:#2E75B6;color:white;
border-radius:20px;padding:.2rem .8rem;font-size:.78rem;font-weight:bold;
margin-bottom:.6rem;}
.next-step{background:#E8F1FB;border:1px dashed #2E75B6;border-radius:10px;
padding:.9rem 1.2rem;margin-top:1rem;color:#1F3864;font-size:.88rem;}
.kpi{background:white;border:1px solid #BDD7EE;border-radius:8px;
padding:.8rem;text-align:center;}
.kpi .v{font-size:1.1rem;font-weight:bold;color:#1F3864;}
.kpi .l{font-size:.72rem;color:#888;margin-top:.3rem;}
.ok{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;
padding:.9rem 1.3rem;color:#375623;margin:.5rem 0;}
.er{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;
padding:.9rem 1.3rem;color:#7B2C00;}
.doc-section{background:#f8fafd;border-radius:10px;padding:1.2rem 1.4rem;
margin-bottom:1rem;border:1px solid #e0eaf5;}
.doc-section h4{color:#1F3864;margin:0 0 .6rem;}
.tag{display:inline-block;background:#D6E4F0;color:#1F3864;
border-radius:4px;padding:.1rem .5rem;font-size:.78rem;
font-family:monospace;margin:.1rem;}
div[data-testid="stDownloadButton"] button{
background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""<div class="hdr">
<h1>📊 FiscalXL Pro</h1>
<p>Suite de notation des entreprises marocaines · Module 1 — Conversion PDF fiscal</p>
</div>""", unsafe_allow_html=True)

# ── Navigation ────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "🏠 Accueil & Conversion",
    "📖 Guide d'utilisation",
    "🔧 Documentation technique"
])


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 1 — ACCUEIL & CONVERSION
# ══════════════════════════════════════════════════════════════════════════════
with tab1:

    # Message de bienvenue
    st.markdown("""<div class="welcome">
<span class="step-badge">Étape 1 / 2</span>
<h3>Bienvenue sur FiscalXL Pro</h3>
<p>
FiscalXL Pro est le premier module de notre suite d'analyse et de notation des entreprises marocaines.
Il convertit automatiquement les bilans fiscaux au format PDF (AMMC et DGI) en fichiers Excel
structurés, conformes au Modèle Comptable Normal (MCN, loi 9-88).
</p>
<div class="next-step">
⏭️ <strong>Prochaine étape :</strong>
Les fichiers Excel générés s'intègrent directement dans la <strong>moulinette d'analyse financière</strong>
pour le calcul des ratios et la notation des entreprises.
</div>
</div>""", unsafe_allow_html=True)

    st.markdown("---")

    # Sidebar format
    with st.sidebar:
        st.markdown("### 📋 Format du PDF")
        fmt_choice = st.radio(
            "",
            ["📄 AMMC — 5 pages", "🏛️ DGI — 7 pages"],
            index=0
        )
        is_dgi    = "DGI" in fmt_choice
        fmt_label = "DGI" if is_dgi else "AMMC"

        if is_dgi:
            st.markdown("""
**Format DGI :**
- 7 pages
- Actif × 2 pages
- Passif × 1 page
- CPC × 3 pages
            """)
        else:
            st.markdown("""
**Format AMMC :**
- 5 pages
- Actif × 1 page
- Passif × 1 page
- CPC × 2 pages
            """)

        st.markdown("---")
        st.markdown("""
**Excel généré — 4 feuilles :**
- `1 - Identification`
- `2 - Bilan Actif`
- `3 - Bilan Passif`
- `4 - CPC`
""")
        st.caption("FiscalXL Pro · MCN loi 9-88 · v1.0")

    # Upload
    pages_info = {
        "AMMC": "5 pages (Actif / Passif / CPC × 2)",
        "DGI":  "7 pages (Actif × 2 / Passif / CPC × 3)",
    }

    st.markdown(f"### 📂 Importer le PDF fiscal")
    uploaded = st.file_uploader(
        f"Format {fmt_label} — {pages_info[fmt_label]}",
        type=["pdf"],
        help="Bilan fiscal complet au format PDF, généré par le logiciel comptable"
    )

    if not uploaded:
        st.markdown(f"""<div style="text-align:center;padding:3rem;color:#888;
border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;">
<div style="font-size:3rem;">📄</div>
<h3 style="color:#2E75B6;">Glissez votre PDF ici</h3>
<p>Format <strong>{fmt_label}</strong> · {pages_info[fmt_label]}</p>
<p style="font-size:.82rem;color:#aaa;">Le fichier est traité localement et non conservé</p>
</div>""", unsafe_allow_html=True)
    else:
        st.markdown("---")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded.getbuffer())
            pdf_path = tmp.name
        output_path = pdf_path.replace(".pdf", "_out.xlsx")

        try:
            progress = st.progress(0)
            status   = st.empty()

            status.info("📄 Lecture du PDF en cours...")
            progress.progress(20)

            if is_dgi:
                from core.dgi_parser import parse
            else:
                from core.ammc_parser import parse
            from core.excel_writer import write

            progress.progress(35)

            with st.spinner("Analyse du PDF..."):
                parsed = parse(pdf_path)

            progress.progress(65)
            status.info("📊 Génération de l'Excel...")

            with st.spinner("Écriture Excel..."):
                stats = write(parsed, output_path)

            progress.progress(100)
            status.empty()

            info     = parsed['info']
            raison   = (info.get('raison_sociale') or '—')[:30]
            exercice = info.get('exercice_fin') or '—'

            # KPIs
            for col, (lbl, val) in zip(st.columns(4), [
                ("Raison Sociale", raison),
                ("Date de bilan",  exercice),
                ("Format",         stats['format']),
                ("Lignes Excel",   f"{stats['rows']} lignes"),
            ]):
                col.markdown(
                    f'<div class="kpi"><div class="v">{val}</div>'
                    f'<div class="l">{lbl}</div></div>',
                    unsafe_allow_html=True
                )

            st.markdown("<br>", unsafe_allow_html=True)

            st.markdown(f"""<div class="ok">
✅ <strong>Conversion réussie</strong>
&nbsp;·&nbsp; Format <strong>{stats['format']}</strong>
&nbsp;·&nbsp; Actif : <strong>{stats['actif']}</strong> lignes
&nbsp;·&nbsp; Passif : <strong>{stats['passif']}</strong> lignes
&nbsp;·&nbsp; CPC : <strong>{stats['cpc']}</strong> lignes
</div>""", unsafe_allow_html=True)

            fname = (
                f"FiscalXL_{raison.replace(' ','_')[:20]}"
                f"_{exercice.replace('/','_')}"
                f"_{stats['format']}.xlsx"
            )
            with open(output_path, "rb") as f:
                st.download_button(
                    "📥 Télécharger l'Excel structuré",
                    data=f,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            logger.exception("Erreur conversion")
            st.markdown(
                f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                unsafe_allow_html=True
            )
            import traceback
            st.code(traceback.format_exc())
        finally:
            for f in [pdf_path, output_path]:
                try:
                    if os.path.exists(f): os.unlink(f)
                except Exception:
                    pass


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 2 — GUIDE D'UTILISATION
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("## 📖 Guide d'utilisation")
    st.markdown("Ce guide explique comment utiliser FiscalXL Pro pour convertir un bilan fiscal PDF en Excel structuré.")

    st.markdown("---")

    # Étape 1
    st.markdown("### Étape 1 — Choisir le format du PDF")
    st.markdown("""
Dans la **barre latérale gauche**, sélectionnez le format de votre PDF avant d'importer :

| Format | Pages | Utilisation |
|--------|-------|-------------|
| **AMMC** | 5 pages | Bilans déposés auprès de l'AMMC — Modèle Normal standard |
| **DGI** | 7 pages | États de synthèse conformes à la déclaration DGI |

> ⚠️ Si vous choisissez le mauvais format, la conversion peut échouer ou produire des valeurs incorrectes.
""")

    # Capture 1
    st.markdown("**📸 Capture 1 — Sélection du format dans la barre latérale**")
    col_img, col_txt = st.columns([1, 1])
    with col_img:
        # Placeholder pour la capture — remplacer par st.image("assets/cap1_format.png")
        st.info("📷 Insérer ici : `st.image('assets/cap1_format.png')`\n\nCapture de la sidebar avec le radio bouton AMMC/DGI sélectionné.")
    with col_txt:
        st.markdown("""
**Comment faire :**
1. Regardez la barre à gauche de l'écran
2. Cliquez sur **AMMC** ou **DGI** selon votre document
3. La description du format s'affiche en dessous
        """)

    st.markdown("---")

    # Étape 2
    st.markdown("### Étape 2 — Importer le PDF")
    st.markdown("""
Dans la zone centrale, **glissez-déposez** votre fichier PDF ou cliquez sur **Browse files**.

Le fichier doit être :
- Au format **PDF** uniquement
- Le bilan fiscal **complet** (toutes les pages)
- Non protégé par mot de passe
""")

    # Capture 2
    st.markdown("**📸 Capture 2 — Zone d'upload du PDF**")
    col_img2, col_txt2 = st.columns([1, 1])
    with col_img2:
        st.info("📷 Insérer ici : `st.image('assets/cap2_upload.png')`\n\nCapture de la zone d'upload avec le cadre pointillé.")
    with col_txt2:
        st.markdown("""
**Comment faire :**
1. Cliquez sur la zone pointillée ou glissez le fichier
2. Sélectionnez votre PDF dans l'explorateur
3. Le nom du fichier s'affiche une fois chargé
4. La conversion démarre automatiquement
        """)

    st.markdown("---")

    # Étape 3
    st.markdown("### Étape 3 — Télécharger l'Excel")
    st.markdown("""
Après la conversion, les **indicateurs clés** s'affichent (raison sociale, date de bilan, nombre de lignes).
Cliquez ensuite sur **Télécharger l'Excel structuré**.
""")

    # Capture 3
    st.markdown("**📸 Capture 3 — Résultat et bouton de téléchargement**")
    col_img3, col_txt3 = st.columns([1, 1])
    with col_img3:
        st.info("📷 Insérer ici : `st.image('assets/cap3_resultat.png')`\n\nCapture des KPIs verts et du bouton de téléchargement bleu.")
    with col_txt3:
        st.markdown("""
**Ce que vous obtenez :**
- Fichier Excel nommé automatiquement
- 4 feuilles : Identification, Actif, Passif, CPC
- Format MCN standardisé
- Prêt pour intégration dans la moulinette
        """)

    st.markdown("---")

    # Étape 4
    st.markdown("### Étape 4 — Structure de l'Excel généré")
    st.markdown("""
Le fichier Excel contient **4 feuilles** dans un format fixe et standardisé :
""")

    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.markdown("""
**1 - Identification**
- Raison sociale
- Identifiant fiscal
- Taxe professionnelle
- Adresse
- Date de bilan
""")
    with col_b:
        st.markdown("""
**2 - Bilan Actif**
- Brut
- Amortissements
- Net Exercice N
- Net Exercice N-1
""")
    with col_c:
        st.markdown("""
**3 - Bilan Passif**
- Exercice N
- Exercice N-1
""")
    with col_d:
        st.markdown("""
**4 - CPC**
- Propres à l'exercice
- Exercices précédents
- Totaux N
- Totaux N-1
""")

    # Capture 4
    st.markdown("**📸 Capture 4 — Exemple de feuille Bilan Actif dans Excel**")
    st.info("📷 Insérer ici : `st.image('assets/cap4_excel_actif.png', use_column_width=True)`\n\nCapture d'écran de la feuille '2 - Bilan Actif' ouverte dans Excel avec les données réelles.")

    st.markdown("---")
    st.info("💡 **Astuce :** Une fois votre Excel généré, passez à l'**Étape 2** : intégrez-le dans la moulinette d'analyse financière pour obtenir les ratios et la notation de l'entreprise.")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 3 — DOCUMENTATION TECHNIQUE
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("## 🔧 Documentation technique")
    st.markdown("Documentation à destination de l'équipe Gestion des Risques.")

    st.markdown("---")

    # Architecture
    st.markdown("### Architecture du projet")
    st.code("""
FiscalXL_Pro/
├── app.py                  ← Interface Streamlit (ce fichier)
├── core/
│   ├── ammc_parser.py      ← Parser PDF format AMMC (5 pages)
│   ├── dgi_parser.py       ← Parser PDF format DGI (7 pages)
│   ├── excel_writer.py     ← Génération Excel (commun AMMC et DGI)
│   └── synonyms.py         ← Dictionnaire des variantes de labels
├── utils/
│   └── logger.py           ← Journalisation
└── requirements.txt        ← Dépendances Python
    """, language="")

    st.markdown("---")

    # Formats supportés
    st.markdown("### Formats PDF supportés")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="doc-section"><h4>📄 Format AMMC</h4>', unsafe_allow_html=True)
        st.markdown("""
**Structure :** 5 pages fixes
- Page 1 : Identification (raison sociale, IF, exercice)
- Page 2 : Bilan Actif
- Page 3 : Bilan Passif
- Pages 4-5 : CPC (Compte de Produits et Charges)

**Extraction :** `core/ammc_parser.py`
Template fixe MCN → matching par synonymes → 49 postes Actif, 44 Passif, 54 CPC
""")
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="doc-section"><h4>🏛️ Format DGI</h4>', unsafe_allow_html=True)
        st.markdown("""
**Structure :** 7 pages fixes
- Page 1 : Identification
- Pages 2-3 : Bilan Actif (2 pages)
- Page 4 : Bilan Passif
- Pages 5-7 : CPC (3 pages)

**Extraction :** `core/dgi_parser.py`
Même template MCN que AMMC → même Excel en sortie
""")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Pipeline
    st.markdown("### Pipeline de traitement")
    st.markdown("""
```
PDF fiscal
    ↓
pdfplumber.extract_tables()        ← Extraction brute des tableaux
    ↓
_detect_val_cols()                 ← Détection automatique des colonnes numériques
    ↓
match_label() + synonyms.py        ← Matching des labels vers le template MCN
    ↓
_build_value_map()                 ← Mapping {template_idx: [v1, v2, v3, v4]}
    ↓
excel_writer.write()               ← Génération Excel 4 feuilles
    ↓
Excel structuré MCN
```
""")

    st.markdown("---")

    # Excel output
    st.markdown("### Structure de l'Excel généré")

    st.markdown('<div class="doc-section"><h4>Colonnes par feuille</h4>', unsafe_allow_html=True)
    st.markdown("""
| Feuille | Col A | Col B | Col C | Col D | Col E |
|---------|-------|-------|-------|-------|-------|
| **2 - Bilan Actif** | Désignation | Brut | Amort. & Prov. | Net N | Net N-1 |
| **3 - Bilan Passif** | Désignation | Exercice N | Exercice N-1 | — | — |
| **4 - CPC** | N° | Désignation | Propres N | Exerc. Préc. | Total N | Total N-1 |
""")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Cas limites
    st.markdown("### Cas limites et comportements connus")

    st.markdown("""
| Situation | Comportement |
|-----------|-------------|
| PDF avec cellules fusionnées (ex : SGTM) | Extraction X/Y automatique — résultats partiels possibles |
| Label non reconnu | Ignoré (non mappé dans le template) |
| Valeur manquante dans le PDF | Affiché comme `0` dans l'Excel |
| PDF protégé par mot de passe | Erreur à l'extraction |
| Moins de pages que prévu | Extraction partielle des sections disponibles |
""")

    st.markdown("---")

    # Dépendances
    st.markdown("### Dépendances Python")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.markdown("""
<span class="tag">streamlit</span>
<span class="tag">pdfplumber</span>
<span class="tag">openpyxl</span>
""", unsafe_allow_html=True)
    with col_d2:
        st.markdown("""
<span class="tag">python >= 3.10</span>
<span class="tag">re</span>
<span class="tag">unicodedata</span>
""", unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("### Contact & support")
    st.markdown("""
Pour tout problème de conversion ou PDF non reconnu, contacter l'équipe technique avec :
- Le fichier PDF concerné
- Le format attendu (AMMC ou DGI)
- Le message d'erreur affiché
""")
