"""FiscalXL — Convertisseur PDF fiscal → Excel structuré (MCN loi 9-88)
Formats supportés :
  • AMMC  — 5 pages (Actif/Passif/CPC×2)
  • DGI   — 7 pages (Actif×2/Passif/CPC×3)
  • SGTM  — 5 pages (Garde/Actif/Passif/CPC×2)

Moulinette — Remplissage automatique à partir de fichiers Excel FiscalXL
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
.mol-card{background:#f8fafd;border:1px solid #BDD7EE;border-radius:10px;
padding:.8rem 1rem;margin:.3rem 0;font-size:.88rem;}
.mol-card .yr{font-weight:bold;color:#1F3864;font-size:1rem;}
.mol-card .col{color:#2E75B6;font-weight:bold;}
div[data-testid="stDownloadButton"] button{
background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr">
<h1>📊 FiscalXL</h1>
<p>Convertisseur PDF → Excel structuré · MCN loi 9-88 Maroc · Structure fixe garantie</p>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# NAVIGATION
# ══════════════════════════════════════════════════════════════════

page = st.radio(
    "",
    ["📄 Convertir PDF → Excel", "📊 Remplir la Moulinette"],
    horizontal=True,
    label_visibility="collapsed",
)
st.markdown("---")

# ══════════════════════════════════════════════════════════════════
# PAGE 1 — CONVERTIR PDF → EXCEL
# ══════════════════════════════════════════════════════════════════

if page == "📄 Convertir PDF → Excel":

    # ── Sidebar ──────────────────────────────────────────────────
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

        is_dgi   = "DGI"  in fmt_choice
        is_sgtm  = "SGTM" in fmt_choice
        fmt_label = "DGI" if is_dgi else ("SGTM" if is_sgtm else "AMMC")

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

    # ── Upload ────────────────────────────────────────────────────
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

        info     = parsed['info']
        raison   = (info.get('raison_sociale') or '—')[:28]
        exercice = info.get('exercice_fin') or '—'

        cols = st.columns(4)
        kpis = [
            ("Raison Sociale", raison),
            ("Fin Exercice",   exercice),
            ("Format",         stats['format']),
            ("Lignes Excel",   f"{stats['rows']} (fixe)"),
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

        if is_sgtm:
            id_fiscal = info.get('identifiant_fiscal', '')
            taxe_pro  = info.get('taxe_pro', '')
            adresse   = info.get('adresse', '')
            if id_fiscal or taxe_pro:
                st.markdown(f"""<div style="background:#f0f4f8;border-radius:8px;
padding:.7rem 1rem;font-size:.85rem;color:#444;margin:.4rem 0;">
🏢 <b>IF :</b> {id_fiscal} &nbsp;|&nbsp; <b>TP :</b> {taxe_pro}
&nbsp;|&nbsp; {adresse}
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
                if os.path.exists(f): os.unlink(f)
            except Exception:
                pass


# ══════════════════════════════════════════════════════════════════
# PAGE 2 — MOULINETTE
# ══════════════════════════════════════════════════════════════════

else:
    st.markdown("### 📊 Remplir la Moulinette")
    st.markdown(
        "Importez **2 ou 3 fichiers Excel FiscalXL** (générés par la page précédente). "
        "L'application détecte automatiquement l'année de chaque fichier et remplit "
        "la moulinette : **D = N-2 · E = N-1 · F = N**."
    )

    uploaded_excels = st.file_uploader(
        "Fichiers Excel FiscalXL",
        type=["xlsx"],
        accept_multiple_files=True,
        help="2 ou 3 fichiers Excel générés par FiscalXL (un par exercice)"
    )

    if not uploaded_excels:
        st.markdown("""<div style="text-align:center;padding:3rem;color:#888;
border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;">
<div style="font-size:3rem;">📊</div>
<h3 style="color:#2E75B6;">Importez vos fichiers Excel FiscalXL</h3>
<p>2 exercices minimum · 3 exercices maximum</p>
<p style="font-size:.85rem;">Les fichiers sont triés automatiquement par année</p>
</div>""", unsafe_allow_html=True)
        st.stop()

    # Vérification du nombre de fichiers
    n = len(uploaded_excels)
    if n > 3:
        st.error(f"❌ Maximum 3 fichiers. Vous avez importé {n} fichiers.")
        st.stop()

    if n == 1:
        st.warning("⚠️ Vous avez importé 1 seul fichier. La moulinette sera remplie sur 1 colonne.")

    # Aperçu des fichiers importés
    st.markdown(f"**{n} fichier(s) importé(s) :**")

    from core.moulinette import read_fiscalxl
    import tempfile

    previews = []
    tmp_paths = []

    for uf in uploaded_excels:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uf.getbuffer())
            tmp_path = tmp.name
        tmp_paths.append(tmp_path)
        try:
            d = read_fiscalxl(tmp_path)
            previews.append(d)
        except Exception as e:
            st.error(f"❌ Erreur lecture {uf.name} : {e}")
            st.stop()

    # Trier par année pour l'aperçu
    previews_sorted = sorted(zip(previews, tmp_paths), key=lambda x: x[0]['annee'])
    col_labels = ['D — N-2', 'E — N-1', 'F — N']

    cols_prev = st.columns(n)
    for i, (d, _) in enumerate(previews_sorted):
        with cols_prev[i]:
            col_lbl = col_labels[i] if i < len(col_labels) else f"Col {i+1}"
            st.markdown(f"""<div class="mol-card">
<div class="col">{col_lbl}</div>
<div class="yr">📅 {d['annee'] or '?'}</div>
<div>{d['societe'][:30] or '—'}</div>
<div style="color:#666;font-size:.8rem;">{d['exercice'] or '—'}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Bouton
    if st.button("📊 Insérer dans la Moulinette", use_container_width=True, type="primary"):
        with st.spinner("Remplissage de la moulinette en cours..."):
            try:
                from core.moulinette import fill_moulinette

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out:
                    out_path = out.name

                sorted_paths = [p for _, p in previews_sorted]
                result = fill_moulinette(sorted_paths, out_path)

                # Résumé
                st.markdown(f"""<div class="ok">
✅ <strong>Moulinette remplie !</strong>
&nbsp;·&nbsp; <strong>{result['filled']}</strong> cellules injectées
&nbsp;·&nbsp; <strong>{len(result['datasets'])}</strong> exercice(s)
</div>""", unsafe_allow_html=True)

                # Détail par exercice
                for d in result['datasets']:
                    st.markdown(f"""<div class="mol-card">
<span class="col">{d['col']}</span> &nbsp;→&nbsp;
<strong>{d['societe'][:35] or '—'}</strong>
&nbsp;|&nbsp; {d['exercice'] or '—'}
</div>""", unsafe_allow_html=True)

                # Téléchargement
                societe_name = result['datasets'][-1]['societe'][:20].replace(' ', '_')
                fname = f"Moulinette_{societe_name}.xlsx"

                with open(out_path, "rb") as f:
                    st.download_button(
                        "📥 Télécharger la Moulinette remplie",
                        data=f,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                os.unlink(out_path)

            except Exception as e:
                logger.exception("Erreur moulinette")
                st.markdown(
                    f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                    unsafe_allow_html=True
                )
                import traceback
                st.code(traceback.format_exc())

    # Nettoyage fichiers temporaires
    for p in tmp_paths:
        try:
            if os.path.exists(p): os.unlink(p)
        except Exception:
            pass
