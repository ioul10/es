"""FiscalXL — Convertisseur PDF fiscal → Excel"""
import streamlit as st
import tempfile, os
from core.pdf_to_excel import convert
from utils.logger import get_logger

logger = get_logger(__name__)

st.set_page_config(page_title="FiscalXL", page_icon="📊", layout="wide")

st.markdown("""<style>
.hdr{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.4rem 2rem;
  border-radius:12px;margin-bottom:1.2rem;}
.hdr h1{color:white;margin:0;font-size:1.8rem;}
.hdr p{color:#BDD7EE;margin:.3rem 0 0;font-size:.88rem;}
.kpi{background:white;border:1px solid #BDD7EE;border-radius:8px;
  padding:.8rem;text-align:center;}
.kpi .v{font-size:1.1rem;font-weight:bold;color:#1F3864;}
.kpi .l{font-size:.72rem;color:#888;margin-top:.3rem;}
.ok{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;
  padding:.9rem 1.3rem;color:#375623;margin:.5rem 0;}
.er{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;
  padding:.9rem 1.3rem;color:#7B2C00;}
div[data-testid="stDownloadButton"] button{
  background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
  border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr">
<h1>📊 FiscalXL</h1>
<p>Convertisseur PDF → Excel · Pièces annexes IS — MCN loi 9-88 Maroc</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Format PDF")
    fmt_choice = st.radio(
        "Sélectionner le format :",
        ["🔵 Détection automatique", "📄 AMMC — 5 pages", "🏛️ DGI — 7 pages"],
        index=0,
    )
    st.markdown("---")
    st.markdown("""
**AMMC — 5 pages**
- Page 1 : Identification
- Page 2 : Bilan Actif
- Page 3 : Bilan Passif
- Pages 4-5 : CPC

**DGI — 7 pages**
- Page 1 : Identification
- Pages 2-3 : Bilan Actif
- Page 4 : Bilan Passif
- Pages 5-7 : CPC
""")
    st.caption("FiscalXL · MCN loi 9-88")

# ── Upload ───────────────────────────────────────────────────────────────────
col1, col2 = st.columns([3, 2])
with col1:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader("Glissez-déposez ou cliquez", type=["pdf"])

with col2:
    st.markdown("### 🔄 Structure générée")
    for sheet, desc in [
        ("1 - Identification", "Raison sociale, IF, exercice"),
        ("2 - Bilan Actif",    "Brut | Amort. | Net N | Net N-1"),
        ("3 - Bilan Passif",   "Exercice N | Exercice N-1"),
        ("4 - CPC",            "Propres N | Préc. | Total N | N-1"),
    ]:
        st.markdown(
            f'<div style="background:#f8f9fa;border-left:4px solid #2E75B6;'
            f'padding:.4rem .8rem;border-radius:0 6px 6px 0;margin:.25rem 0;">'
            f'<strong style="color:#1F3864;font-size:.82rem">{sheet}</strong><br>'
            f'<span style="color:#666;font-size:.75rem">{desc}</span></div>',
            unsafe_allow_html=True)

if not uploaded:
    st.markdown("""<div style="text-align:center;padding:3rem;color:#888;
      border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;margin-top:1rem;">
      <div style="font-size:3rem;">📄</div>
      <h3 style="color:#2E75B6;">Importez un PDF pour commencer</h3>
      <p>Formats supportés : AMMC (5 pages) et DGI (7 pages)</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

st.markdown("---")

# ── Conversion ───────────────────────────────────────────────────────────────
with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
    tmp.write(uploaded.getbuffer())
    pdf_path = tmp.name
output_path = pdf_path.replace(".pdf", "_out.xlsx")

try:
    progress = st.progress(0)
    status   = st.empty()

    status.info("📄 Lecture du PDF...")
    progress.progress(20)

    with st.spinner("Conversion en cours..."):
        stats = convert(pdf_path, output_path)

    progress.progress(100)
    status.empty()

    info    = stats['info']
    raison  = (info.get('raison_sociale') or '—')[:30]
    exercice = info.get('exercice_fin') or '—'
    fmt_det  = stats.get('format', '?')

    # KPIs
    cols = st.columns(4)
    for col, (lbl, val) in zip(cols, [
        ("Raison Sociale", raison),
        ("Fin exercice",   exercice),
        ("Format détecté", fmt_det),
        ("Lignes Excel",   str(stats['rows'])),
    ]):
        col.markdown(f'<div class="kpi"><div class="v">{val}</div>'
                     f'<div class="l">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"""<div class="ok">
      ✅ <strong>Conversion réussie !</strong>
      &nbsp;·&nbsp; Format <strong>{fmt_det}</strong>
      &nbsp;·&nbsp; {stats['rows']} lignes extraites
      &nbsp;·&nbsp; 4 feuilles Excel générées
    </div>""", unsafe_allow_html=True)

    # Download
    fname = (f"{raison.replace(' ','_')[:20]}_"
             f"{exercice.replace('/','_')}_fiscal.xlsx")
    with open(output_path, "rb") as f:
        st.download_button(
            "📥 Télécharger le fichier Excel",
            data=f, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except Exception as e:
    logger.exception("Erreur pipeline")
    st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                unsafe_allow_html=True)
    import traceback; st.code(traceback.format_exc())
finally:
    for f in [pdf_path, output_path]:
        try:
            if os.path.exists(f): os.unlink(f)
        except: pass
