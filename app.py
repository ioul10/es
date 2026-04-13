"""FiscalXL — Convertisseur PDF fiscal → Excel + Moulinette automatique"""
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
.mol-card{background:#f8fafd;border:1px solid #BDD7EE;border-radius:10px;
padding:.8rem 1rem;margin:.3rem 0;font-size:.88rem;}
.mol-card .yr{font-weight:bold;color:#1F3864;font-size:1rem;}
.mol-card .col{color:#2E75B6;font-weight:bold;font-size:.9rem;}
div[data-testid="stDownloadButton"] button{
background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr">
<h1>📊 FiscalXL</h1>
<p>Convertisseur PDF → Excel structuré · MCN loi 9-88 Maroc · Moulinette automatique</p>
</div>""", unsafe_allow_html=True)

# ── Navigation ────────────────────────────────────────────────────────────────
page = st.radio("", ["📄 Convertir PDF → Excel", "📊 Remplir la Moulinette"],
                horizontal=True, label_visibility="collapsed")
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — CONVERTIR PDF → EXCEL
# ══════════════════════════════════════════════════════════════════════════════
if page == "📄 Convertir PDF → Excel":

    with st.sidebar:
        st.markdown("### 📋 Format PDF")
        fmt_choice = st.radio("", ["📄 AMMC — 5 pages", "🏛️ DGI — 7 pages",
                                    "🏗️ SGTM — 5 pages (Modèle IS)"], index=0)
        is_dgi   = "DGI"  in fmt_choice
        is_sgtm  = "SGTM" in fmt_choice
        fmt_label = "DGI" if is_dgi else ("SGTM" if is_sgtm else "AMMC")
        if is_sgtm:
            st.markdown("**Format SGTM :**\n- Page 1 : Garde\n- Page 2 : Actif\n- Page 3 : Passif\n- Pages 4-5 : CPC")
        elif is_dgi:
            st.markdown("**Format DGI :**\n- 7 pages : Actif×2 / Passif / CPC×3")
        else:
            st.markdown("**Format AMMC :**\n- 5 pages : Actif / Passif / CPC×2")
        st.markdown("---")
        st.markdown(f"**Excel généré :**\n- `1 - Identification`\n- `2 - Bilan Actif`\n- `3 - Bilan Passif`\n- `4 - CPC`\n\n_Format {fmt_label}_")
        st.caption("FiscalXL · MCN loi 9-88")

    st.markdown("### 📂 Importer le PDF")
    pages_info = {"AMMC":"5 pages (Actif/Passif/CPC×2)",
                  "DGI":"7 pages (Actif×2/Passif/CPC×3)",
                  "SGTM":"5 pages (Garde/Actif/Passif/CPC×2)"}

    uploaded = st.file_uploader(f"PDF {fmt_label} ({pages_info[fmt_label]})", type=["pdf"])

    if not uploaded:
        st.markdown(f"""<div style="text-align:center;padding:3rem;color:#888;
border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;">
<div style="font-size:3rem;">📄</div>
<h3 style="color:#2E75B6;">Importez un PDF {fmt_label}</h3>
<p>{pages_info[fmt_label]}</p></div>""", unsafe_allow_html=True)
        st.stop()

    st.markdown("---")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded.getbuffer()); pdf_path = tmp.name
    output_path = pdf_path.replace(".pdf","_out.xlsx")

    try:
        progress = st.progress(0); status = st.empty()
        status.info("📄 Extraction en cours..."); progress.progress(20)

        if is_sgtm:
            from core.sgtm_parser import parse; from core.sgtm_excel_writer import write
        elif is_dgi:
            from core.dgi_parser import parse; from core.excel_writer import write
        else:
            from core.ammc_parser import parse; from core.excel_writer import write

        progress.progress(35)
        with st.spinner("Analyse du PDF..."): parsed = parse(pdf_path)
        progress.progress(65); status.info("📊 Génération Excel...")
        with st.spinner("Écriture Excel..."): stats = write(parsed, output_path)
        progress.progress(100); status.empty()

        info = parsed['info']
        raison = (info.get('raison_sociale') or '—')[:28]
        exercice = info.get('exercice_fin') or '—'

        for col,(lbl,val) in zip(st.columns(4),
            [("Raison Sociale",raison),("Fin Exercice",exercice),
             ("Format",stats['format']),("Lignes Excel",f"{stats['rows']} (fixe)")]):
            col.markdown(f'<div class="kpi"><div class="v">{val}</div>'
                         f'<div class="l">{lbl}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""<div class="ok">✅ <strong>Excel généré !</strong>
&nbsp;·&nbsp; Format <strong>{stats['format']}</strong>
&nbsp;·&nbsp; <strong>{stats['actif']}</strong> lignes Actif
&nbsp;·&nbsp; <strong>{stats['passif']}</strong> lignes Passif
&nbsp;·&nbsp; <strong>{stats['cpc']}</strong> lignes CPC</div>""", unsafe_allow_html=True)

        if is_sgtm:
            id_f = info.get('identifiant_fiscal',''); tp = info.get('taxe_pro','')
            if id_f or tp:
                st.markdown(f'<div style="background:#f0f4f8;border-radius:8px;'
                            f'padding:.7rem 1rem;font-size:.85rem;color:#444;margin:.4rem 0;">'
                            f'🏢 <b>IF :</b> {id_f} &nbsp;|&nbsp; <b>TP :</b> {tp}</div>',
                            unsafe_allow_html=True)

        fname = f"{raison.replace(' ','_')[:20]}_{exercice.replace('/','_')}_{stats['format']}.xlsx"
        with open(output_path,"rb") as f:
            st.download_button("📥 Télécharger le fichier Excel", data=f, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        logger.exception("Erreur FiscalXL")
        st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                    unsafe_allow_html=True)
        import traceback; st.code(traceback.format_exc())
    finally:
        for f in [pdf_path, output_path]:
            try:
                if os.path.exists(f): os.unlink(f)
            except: pass

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — MOULINETTE
# ══════════════════════════════════════════════════════════════════════════════
else:
    st.markdown("### 📊 Remplir la Moulinette")
    st.markdown(
        "Importez **2 fichiers Excel FiscalXL** : l'année **N** (récente) et l'année **N-1**. "
        "La colonne N-2 est extraite automatiquement depuis l'exercice précédent du fichier N-1."
    )

    cols_info = st.columns(3)
    with cols_info[0]:
        st.markdown('<div class="mol-card"><div class="col">D — N-2</div>'
                    '<div style="color:#666;font-size:.8rem;">Exercice précédent<br>du fichier N-1</div></div>',
                    unsafe_allow_html=True)
    with cols_info[1]:
        st.markdown('<div class="mol-card"><div class="col">E — N-1</div>'
                    '<div style="color:#666;font-size:.8rem;">Exercice N<br>du fichier N-1</div></div>',
                    unsafe_allow_html=True)
    with cols_info[2]:
        st.markdown('<div class="mol-card"><div class="col">F — N</div>'
                    '<div style="color:#666;font-size:.8rem;">Exercice N<br>du fichier N</div></div>',
                    unsafe_allow_html=True)

    st.markdown("---")

    col_n, col_n1 = st.columns(2)
    with col_n:
        st.markdown("#### 📄 Fichier N — Année récente")
        upload_n = st.file_uploader("Excel FiscalXL — Exercice N", type=["xlsx"],
                                     key="upload_n",
                                     help="Excel de l'année la plus récente")
    with col_n1:
        st.markdown("#### 📄 Fichier N-1 — Année précédente")
        upload_n1 = st.file_uploader("Excel FiscalXL — Exercice N-1", type=["xlsx"],
                                      key="upload_n1",
                                      help="Excel N-1 (contient aussi N-2 en colonne précédente)")

    tmp_n = tmp_n1 = None

    if upload_n or upload_n1:
        st.markdown("---")
        from core.moulinette import read_fiscalxl

        if upload_n:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                f.write(upload_n.getbuffer()); tmp_n = f.name
            info_n = read_fiscalxl(tmp_n)
        else:
            info_n = None

        if upload_n1:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                f.write(upload_n1.getbuffer()); tmp_n1 = f.name
            info_n1 = read_fiscalxl(tmp_n1)
        else:
            info_n1 = None

        st.markdown("**Aperçu des exercices :**")
        prev_cols = st.columns(3)

        with prev_cols[0]:  # D — N-2
            if info_n1:
                an2 = info_n1['annee'] - 1
                st.markdown(f'<div class="mol-card"><div class="col">D — N-2</div>'
                            f'<div class="yr">📅 {an2}</div>'
                            f'<div style="color:#666;font-size:.8rem;">{info_n1["societe"][:28] or "—"}</div>'
                            f'<div style="color:#aaa;font-size:.75rem;">Exercice précédent N-1</div></div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="mol-card" style="opacity:.4"><div class="col">D — N-2</div>'
                            '<div style="color:#999;">En attente fichier N-1</div></div>',
                            unsafe_allow_html=True)

        with prev_cols[1]:  # E — N-1
            if info_n1:
                st.markdown(f'<div class="mol-card"><div class="col">E — N-1</div>'
                            f'<div class="yr">📅 {info_n1["annee"] or "?"}</div>'
                            f'<div>{info_n1["societe"][:28] or "—"}</div>'
                            f'<div style="color:#666;font-size:.8rem;">{info_n1["exercice"] or "—"}</div></div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="mol-card" style="opacity:.4"><div class="col">E — N-1</div>'
                            '<div style="color:#999;">En attente fichier N-1</div></div>',
                            unsafe_allow_html=True)

        with prev_cols[2]:  # F — N
            if info_n:
                st.markdown(f'<div class="mol-card"><div class="col">F — N</div>'
                            f'<div class="yr">📅 {info_n["annee"] or "?"}</div>'
                            f'<div>{info_n["societe"][:28] or "—"}</div>'
                            f'<div style="color:#666;font-size:.8rem;">{info_n["exercice"] or "—"}</div></div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="mol-card" style="opacity:.4"><div class="col">F — N</div>'
                            '<div style="color:#999;">En attente fichier N</div></div>',
                            unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if not (upload_n and upload_n1):
            st.info("⚠️ Importez les 2 fichiers pour activer le remplissage.")
        else:
            if st.button("📊 Insérer dans la Moulinette", use_container_width=True, type="primary"):
                with st.spinner("Remplissage en cours..."):
                    try:
                        from core.moulinette import fill_moulinette
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                            out_path = f.name

                        result = fill_moulinette(tmp_n, tmp_n1, out_path)

                        st.markdown(f"""<div class="ok">
✅ <strong>Moulinette remplie !</strong>
&nbsp;·&nbsp; <strong>{result['filled']}</strong> cellules injectées<br>
D = <strong>{result['annee_n2']}</strong> &nbsp;|&nbsp;
E = <strong>{result['annee_n1']}</strong> &nbsp;|&nbsp;
F = <strong>{result['annee_n']}</strong>
</div>""", unsafe_allow_html=True)

                        societe = result['societe'][:20].replace(' ','_')
                        fname   = f"Moulinette_{societe}_{result['annee_n']}.xlsx"

                        with open(out_path,"rb") as f:
                            st.download_button(
                                "📥 Télécharger la Moulinette remplie",
                                data=f, file_name=fname,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        os.unlink(out_path)

                    except Exception as e:
                        logger.exception("Erreur moulinette")
                        st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                                    unsafe_allow_html=True)
                        import traceback; st.code(traceback.format_exc())

        for p in [tmp_n, tmp_n1]:
            try:
                if p and os.path.exists(p): os.unlink(p)
            except: pass

    else:
        st.markdown("""<div style="text-align:center;padding:3rem;color:#888;
border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;">
<div style="font-size:3rem;">📊</div>
<h3 style="color:#2E75B6;">Importez 2 fichiers Excel FiscalXL</h3>
<p>📄 Fichier N (année récente) &nbsp;+&nbsp; 📄 Fichier N-1 (année précédente)</p>
<p style="font-size:.85rem;color:#aaa;">La colonne N-2 est extraite automatiquement depuis le fichier N-1</p>
</div>""", unsafe_allow_html=True)
