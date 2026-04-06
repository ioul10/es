"""FiscalXL — Convertisseur PDF fiscal → Excel structuré (MCN loi 9-88)
Formats supportés :
  • AMMC  — 5 pages (Actif/Passif/CPC×2)
  • DGI   — 7 pages (Actif×2/Passif/CPC×3)
  • SGTM  — 5 pages (Garde/Actif/Passif/CPC×2)
"""

import streamlit as st
import tempfile, os
from utils.logger import get_logger

logger = get_logger(__name__)

st.set_page_config(page_title="FiscalXL", page_icon="📊", layout="wide")

st.markdown("""<style>
.hdr{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.4rem 2rem;
border-radius:12px;margin-bottom:1.2rem;}
.hdr h1{color:white;margin:0;font-size:1.8rem;}
.hdr p{color:#BDD7EE;margin:.3rem 0 0;}
.kpi{background:white;border:1px solid #BDD7EE;border-radius:8px;
padding:.8rem;text-align:center;}
.kpi .v{font-size:1.1rem;font-weight:bold;color:#1F3864;}
.kpi .l{font-size:.72rem;color:#888;margin-top:.3rem;}
.ok{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;
padding:.9rem 1.3rem;color:#375623;margin:.5rem 0;}
.er{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;
padding:.9rem 1.3rem;color:#7B2C00;}
.fmt-card{border:2px solid #BDD7EE;border-radius:10px;padding:.9rem 1rem;
background:#f8fafd;margin:.4rem 0;cursor:pointer;}
.fmt-card.active{border-color:#2E75B6;background:#EBF3FB;}
div[data-testid="stDownloadButton"] button{
background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr">
<h1>📊 FiscalXL</h1>
<p>Convertisseur PDF → Excel structuré · MCN loi 9-88 Maroc · Structure fixe garantie</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Format PDF")

    fmt_choice = st.radio(
        "",
        [
            "📄 AMMC — 5 pages",
            "🏛️ DGI — 7 pages",
            "🏗️ SGTM — 5 pages (Modèle IS)"
        ],
        index=0
    )

    is_dgi  = "DGI"  in fmt_choice
    is_sgtm = "SGTM" in fmt_choice
    fmt_label = "DGI" if is_dgi else ("SGTM" if is_sgtm else "AMMC")

    # Description du format sélectionné
    if is_sgtm:
        st.markdown("""
**Format SGTM :**
- `Page 1` : Page de garde (identification)
- `Page 2` : Bilan Actif (Tableau 01 1/2)
- `Page 3` : Bilan Passif (Tableau 01 2/2)
- `Page 4` : CPC (Tableau 02 1/2)
- `Page 5` : CPC suite (Tableau 02 2/2)
        """)
    elif is_dgi:
        st.markdown("""
**Format DGI :**
- 7 pages : Actif×2 / Passif / CPC×3
        """)
    else:
        st.markdown("""
**Format AMMC :**
- 5 pages : Actif / Passif / CPC×2
        """)

    st.markdown("---")
    st.markdown(f"""
**Structure Excel générée :**
- `1 - Identification`
- `2 - Bilan Actif`
- `3 - Bilan Passif`
- `4 - CPC`

_Format {fmt_label}_
""")
    st.caption("FiscalXL · MCN loi 9-88")

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Importer le PDF")

pages_info = {
    "AMMC": "5 pages (Actif/Passif/CPC×2)",
    "DGI":  "7 pages (Actif×2/Passif/CPC×3)",
    "SGTM": "5 pages (Garde/Actif/Passif/CPC×2)",
}

uploaded = st.file_uploader(
    f"PDF {fmt_label} ({pages_info[fmt_label]})",
    type=["pdf"]
)

if not uploaded:
    st.markdown(f"""<div style="text-align:center;padding:3rem;color:#888;
border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;">
<div style="font-size:3rem;">📄</div>
<h3 style="color:#2E75B6;">Importez un PDF {fmt_label}</h3>
<p>Format attendu : {pages_info[fmt_label]}</p>
</div>""", unsafe_allow_html=True)
    st.stop()

st.markdown("---")

with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
    tmp.write(uploaded.getbuffer())
    pdf_path = tmp.name

output_path = pdf_path.replace(".pdf", "_out.xlsx")

try:
    progress = st.progress(0)
    status   = st.empty()

    status.info("📄 Extraction en cours...")
    progress.progress(20)

    # ── Chargement du parser selon le format ─────────────────────────────────
    if is_sgtm:
        from core.sgtm_parser import parse
        from core.sgtm_excel_writer import write
    elif is_dgi:
        from core.dgi_parser import parse
        from core.excel_writer import write
    else:
        from core.ammc_parser import parse
        from core.excel_writer import write

    progress.progress(35)

    with st.spinner("Analyse du PDF..."):
        parsed = parse(pdf_path)

    progress.progress(65)
    status.info("📊 Génération Excel...")

    with st.spinner("Écriture Excel..."):
        stats = write(parsed, output_path)

    progress.progress(100)
    status.empty()

    # ── Affichage des résultats ───────────────────────────────────────────────
    info     = parsed['info']
    raison   = (info.get('raison_sociale') or '—')[:28]
    exercice = info.get('exercice_fin') or '—'

    cols = st.columns(4)
    kpis = [
        ("Raison Sociale",   raison),
        ("Fin Exercice",     exercice),
        ("Format",           stats['format']),
        ("Lignes Excel",     f"{stats['rows']} (fixe)"),
    ]
    for col, (lbl, val) in zip(cols, kpis):
        col.markdown(
            f'<div class="kpi"><div class="v">{val}</div>'
            f'<div class="l">{lbl}</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(f"""<div class="ok">
✅ <strong>Excel généré !</strong>
&nbsp;·&nbsp; Format <strong>{stats['format']}</strong>
&nbsp;·&nbsp; <strong>{stats['actif']}</strong> lignes Actif
&nbsp;·&nbsp; <strong>{stats['passif']}</strong> lignes Passif
&nbsp;·&nbsp; <strong>{stats['cpc']}</strong> lignes CPC
</div>""", unsafe_allow_html=True)

    # Informations supplémentaires pour SGTM
    if is_sgtm:
        id_fiscal = info.get('identifiant_fiscal', '')
        taxe_pro  = info.get('taxe_pro', '')
        adresse   = info.get('adresse', '')
        if id_fiscal or taxe_pro:
            st.markdown(f"""<div style="background:#f0f4f8;border-radius:8px;padding:.7rem 1rem;
font-size:.85rem;color:#444;margin:.4rem 0;">
🏢 <b>IF :</b> {id_fiscal} &nbsp;|&nbsp; <b>TP :</b> {taxe_pro} &nbsp;|&nbsp; {adresse}
</div>""", unsafe_allow_html=True)

    fname = (
        f"{raison.replace(' ','_')[:20]}"
        f"_{exercice.replace('/','_')}"
        f"_{stats['format']}.xlsx"
    )

    with open(output_path, "rb") as f:
        st.download_button(
            "📥 Télécharger le fichier Excel",
            data=f,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

except Exception as e:
    logger.exception("Erreur FiscalXL")
    st.markdown(
        f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
        unsafe_allow_html=True
    )
    import traceback
    st.code(traceback.format_exc())

finally:
    for f in [pdf_path, output_path]:
        try:
            if os.path.exists(f):
                os.unlink(f)
        except Exception:
            pass
