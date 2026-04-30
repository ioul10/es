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
.step-badge-rapport{display:inline-block;background:#375623;color:white;
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
.warn{background:#FFF2CC;border:1px solid #FFD966;border-radius:8px;
padding:.9rem 1.3rem;color:#7F6000;margin:.5rem 0;}
.doc-section{background:#f8fafd;border-radius:10px;padding:1.2rem 1.4rem;
margin-bottom:1rem;border:1px solid #e0eaf5;}
.doc-section h4{color:#1F3864;margin:0 0 .6rem;}
.tag{display:inline-block;background:#D6E4F0;color:#1F3864;
border-radius:4px;padding:.1rem .5rem;font-size:.78rem;
font-family:monospace;margin:.1rem;}
.rapport-box{background:#f0f7f0;border-left:4px solid #70AD47;
border-radius:0 10px 10px 0;padding:1.2rem 1.6rem;margin-bottom:1.2rem;}
.rapport-box h3{color:#375623;margin:0 0 .5rem;}
.rapport-box p{color:#444;margin:0;line-height:1.6;}
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
# SIDEBAR — Format
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📋 Format du PDF")
    fmt_choice = st.radio(
        "",
        ["📄 AMMC — 5 pages", "🏛️ DGI — 7 pages", "📑 Rapport financier"],
        index=0
    )
    is_dgi     = "DGI" in fmt_choice
    is_rapport = "Rapport" in fmt_choice
    is_ammc    = not is_dgi and not is_rapport
    fmt_label  = "DGI" if is_dgi else ("Rapport" if is_rapport else "AMMC")

    if is_dgi:
        st.markdown("""
**Format DGI :**
- 7 pages
- Actif × 2 pages
- Passif × 1 page
- CPC × 3 pages
        """)
    elif is_rapport:
        st.markdown("""
**Rapport financier :**
- Pages libres
- Vous indiquez manuellement
  les pages Actif / Passif / CPC
- Identification saisie à la main
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


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 1 — ACCUEIL & CONVERSION
# ══════════════════════════════════════════════════════════════════════════════
with tab1:

    # ── MODE AMMC / DGI ───────────────────────────────────────────────────────
    if not is_rapport:

        st.markdown(f"""<div class="welcome">
<span class="step-badge">Étape 1 / 2</span>
<h3>Bienvenue sur FiscalXL Pro</h3>
<p>
FiscalXL Pro convertit automatiquement les bilans fiscaux au format PDF
(AMMC et DGI) en fichiers Excel structurés, conformes au Modèle Comptable Normal (MCN, loi 9-88).
</p>
<div class="next-step">
⏭️ <strong>Prochaine étape :</strong>
Les fichiers Excel générés s'intègrent directement dans la <strong>moulinette d'analyse financière</strong>
pour le calcul des ratios et la notation des entreprises.
</div>
</div>""", unsafe_allow_html=True)

        st.markdown("---")

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

                info = parsed['info']

                st.markdown(f"""<div class="ok">
✅ <strong>Conversion réussie</strong>
&nbsp;·&nbsp; Format <strong>{stats['format']}</strong>
&nbsp;·&nbsp; Actif : <strong>{stats['actif']}</strong> lignes
&nbsp;·&nbsp; Passif : <strong>{stats['passif']}</strong> lignes
&nbsp;·&nbsp; CPC : <strong>{stats['cpc']}</strong> lignes
</div>""", unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("### ✏️ Vérification et complétion des informations")
                st.markdown(
                    "Vérifiez les informations extraites et complétez les champs manquants. "
                    "**Tous les champs marqués * sont obligatoires.**"
                )

                with st.form("form_validation"):
                    st.markdown("**📋 Informations fiscales**")
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        raison_input = st.text_input("Raison sociale *",
                            value=info.get('raison_sociale', ''),
                            placeholder="Ex : SOCIÉTÉ MAROCAINE DE ...")
                        if_input = st.text_input("Identifiant fiscal *",
                            value=info.get('identifiant_fiscal', ''),
                            placeholder="Ex : 4510887")
                    with col_f2:
                        date_input = st.text_input("Date de bilan *",
                            value=info.get('exercice_fin', '') or info.get('exercice', ''),
                            placeholder="Ex : 31/12/2024")
                        taxe_input = st.text_input("Taxe professionnelle",
                            value=info.get('taxe_professionnelle', ''),
                            placeholder="Ex : 25725940")

                    st.markdown("**🏢 Informations commerciales**")
                    col_c1, col_c2 = st.columns(2)
                    with col_c1:
                        centre_input = st.text_input("Centre d'affaires *",
                            value=info.get('centre_affaires', ''),
                            placeholder="Ex : Casa Finance, Rabat Centre...")
                    with col_c2:
                        secteur_input = st.selectbox("Macro-secteur d'activité *",
                            options=["— Sélectionner —", "Manufactures", "Services",
                                     "Commerce", "BTP", "Holding"])

                    submitted = st.form_submit_button(
                        "📥 Générer et télécharger l'Excel",
                        use_container_width=True, type="primary")

                if submitted:
                    errors = []
                    if not raison_input.strip():  errors.append("Raison sociale")
                    if not if_input.strip():      errors.append("Identifiant fiscal")
                    if not date_input.strip():    errors.append("Date de bilan")
                    if not centre_input.strip():  errors.append("Centre d'affaires")
                    if secteur_input == "— Sélectionner —":
                        errors.append("Macro-secteur d'activité")

                    if errors:
                        st.markdown(
                            f'<div class="er">❌ <strong>Champs obligatoires manquants :</strong> '
                            f'{", ".join(errors)}</div>', unsafe_allow_html=True)
                    else:
                        info.update({
                            'raison_sociale':      raison_input.strip(),
                            'identifiant_fiscal':   if_input.strip(),
                            'exercice_fin':          date_input.strip(),
                            'taxe_professionnelle': taxe_input.strip(),
                            'centre_affaires':       centre_input.strip(),
                            'macro_secteur':         secteur_input,
                        })
                        parsed['info'] = info

                        with st.spinner("Génération de l'Excel..."):
                            stats2 = write(parsed, output_path)

                        fname = (
                            f"FiscalXL_{raison_input.strip().replace(' ','_')[:20]}"
                            f"_{date_input.strip().replace('/','_')}"
                            f"_{stats2['format']}.xlsx"
                        )
                        with open(output_path, "rb") as f_dl:
                            st.download_button(
                                "📥 Télécharger l'Excel", data=f_dl,
                                file_name=fname,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_final")
                        st.markdown(f"""<div class="ok">
✅ <strong>Excel prêt</strong> &nbsp;·&nbsp;
<strong>{raison_input.strip()[:30]}</strong> &nbsp;·&nbsp;
{date_input.strip()} &nbsp;·&nbsp; {secteur_input}
</div>""", unsafe_allow_html=True)

            except Exception as e:
                logger.exception("Erreur conversion")
                st.markdown(
                    f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                    unsafe_allow_html=True)
                import traceback
                st.code(traceback.format_exc())
            finally:
                for f in [pdf_path, output_path]:
                    try:
                        if os.path.exists(f): os.unlink(f)
                    except Exception:
                        pass

    # ══════════════════════════════════════════════════════════════
    # MODE RAPPORT FINANCIER
    # ══════════════════════════════════════════════════════════════
    else:
        st.markdown("""<div class="rapport-box">
<span class="step-badge-rapport">Rapport financier</span>
<h3>Conversion d'un rapport financier libre</h3>
<p>
Ce mode vous permet de traiter des rapports financiers annuels dont la mise en page est libre
(rapports LabelVie, ONCF, etc.). Vous indiquez manuellement les informations d'identification
et les pages contenant le Bilan Actif, le Bilan Passif et le CPC.
</p>
</div>""", unsafe_allow_html=True)

        st.markdown("---")

        # ── ÉTAPE 1 : Upload ─────────────────────────────────────
        st.markdown("### 📂 Étape 1 — Importer le rapport PDF")
        uploaded_r = st.file_uploader(
            "Rapport financier annuel (PDF)",
            type=["pdf"],
            help="Rapport financier complet — toutes les pages",
            key="upload_rapport"
        )

        if not uploaded_r:
            st.markdown("""<div style="text-align:center;padding:2.5rem;color:#888;
border:2px dashed #70AD47;border-radius:12px;background:#f0f7f0;">
<div style="font-size:3rem;">📑</div>
<h3 style="color:#375623;">Glissez votre rapport PDF ici</h3>
<p>Rapport financier annuel · Pages libres</p>
<p style="font-size:.82rem;color:#aaa;">Le fichier est traité localement et non conservé</p>
</div>""", unsafe_allow_html=True)

        else:
            # Sauvegarder le PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_r:
                tmp_r.write(uploaded_r.getbuffer())
                pdf_path_r = tmp_r.name

            # Compter les pages
            try:
                import pdfplumber
                with pdfplumber.open(pdf_path_r) as _pdf:
                    total_pages = len(_pdf.pages)
                st.success(f"✅ PDF chargé — **{total_pages} pages** détectées")
            except Exception as e:
                st.error(f"Erreur lecture PDF : {e}")
                total_pages = 0

            if total_pages > 0:
                st.markdown("---")

                # ── ÉTAPE 2 : Identification manuelle ────────────
                st.markdown("### ✏️ Étape 2 — Informations d'identification")
                st.caption("Ces informations apparaîtront dans la feuille '1 - Identification' de l'Excel.")

                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    r_raison  = st.text_input("Raison sociale *",
                        placeholder="Ex : LABELVIE SA", key="r_raison")
                    r_if      = st.text_input("Identifiant fiscal",
                        placeholder="Ex : 4510887", key="r_if")
                    r_taxe    = st.text_input("Taxe professionnelle",
                        placeholder="Ex : 25725940", key="r_taxe")
                with col_r2:
                    r_date    = st.text_input("Date de bilan *",
                        placeholder="Ex : 31/12/2025", key="r_date")
                    r_centre  = st.text_input("Centre d'affaires *",
                        placeholder="Ex : Casa Finance", key="r_centre")
                    r_secteur = st.selectbox("Macro-secteur *",
                        options=["— Sélectionner —", "Manufactures", "Services",
                                 "Commerce", "BTP", "Holding"],
                        key="r_secteur")

                st.markdown("---")

                # ── ÉTAPE 3 : Sélection des pages ────────────────
                st.markdown("### 📄 Étape 3 — Indiquer les pages")
                st.caption(
                    f"Le rapport contient **{total_pages} pages**. "
                    "Indiquez les numéros de pages (commence à 1). "
                    "Exemples : `1` · `1,2` · `2-4`"
                )

                col_p1, col_p2, col_p3 = st.columns(3)
                with col_p1:
                    st.markdown("**📊 Bilan Actif**")
                    p_actif  = st.text_input("Pages Actif",
                        placeholder="Ex : 1", key="p_actif")
                    st.caption("4 colonnes : Brut / Amort / Net N / Net N-1")
                with col_p2:
                    st.markdown("**📊 Bilan Passif**")
                    p_passif = st.text_input("Pages Passif",
                        placeholder="Ex : 1", key="p_passif")
                    st.caption("2 colonnes : Exercice N / Exercice N-1")
                with col_p3:
                    st.markdown("**📊 CPC**")
                    p_cpc    = st.text_input("Pages CPC",
                        placeholder="Ex : 2,3", key="p_cpc")
                    st.caption("4 colonnes : Propres / Préc / Total N / Total N-1")

                st.markdown("---")

                # ── ÉTAPE 4 : Analyser ────────────────────────────
                st.markdown("### 🔍 Étape 4 — Analyser")

                if st.button("🔍 Analyser le rapport", type="primary",
                             use_container_width=True, key="btn_analyser"):

                    # Validation pages
                    if not p_actif.strip() or not p_passif.strip() or not p_cpc.strip():
                        st.markdown('<div class="er">❌ Veuillez renseigner les pages Actif, Passif et CPC.</div>',
                                    unsafe_allow_html=True)
                    else:
                        try:
                            from core.rapport_parser import parse as parse_rapport

                            info_r = {
                                'raison_sociale':      r_raison.strip(),
                                'identifiant_fiscal':   r_if.strip(),
                                'taxe_professionnelle': r_taxe.strip(),
                                'exercice_fin':          r_date.strip(),
                                'exercice':              r_date.strip(),
                                'adresse':               '',
                                'centre_affaires':       r_centre.strip(),
                                'macro_secteur':         r_secteur,
                                'format':                'Rapport',
                            }

                            with st.spinner("Analyse en cours..."):
                                parsed_r = parse_rapport(
                                    pdf_path_r,
                                    pages_actif=p_actif,
                                    pages_passif=p_passif,
                                    pages_cpc=p_cpc,
                                    info=info_r,
                                )

                            s = parsed_r['_stats']

                            # KPIs résultat
                            st.markdown("#### Résultat de l'analyse")
                            ka, kp, kc = st.columns(3)
                            with ka:
                                pct_a = round(s['actif'] / s['actif_max'] * 100)
                                color_a = "#375623" if pct_a >= 60 else ("#7F6000" if pct_a >= 30 else "#7B2C00")
                                st.markdown(f"""<div class="kpi">
<div class="v" style="color:{color_a};">{s['actif']} / {s['actif_max']}</div>
<div class="l">Postes Actif ({pct_a}%)</div></div>""", unsafe_allow_html=True)
                            with kp:
                                pct_p = round(s['passif'] / s['passif_max'] * 100)
                                color_p = "#375623" if pct_p >= 60 else ("#7F6000" if pct_p >= 30 else "#7B2C00")
                                st.markdown(f"""<div class="kpi">
<div class="v" style="color:{color_p};">{s['passif']} / {s['passif_max']}</div>
<div class="l">Postes Passif ({pct_p}%)</div></div>""", unsafe_allow_html=True)
                            with kc:
                                pct_c = round(s['cpc'] / s['cpc_max'] * 100)
                                color_c = "#375623" if pct_c >= 60 else ("#7F6000" if pct_c >= 30 else "#7B2C00")
                                st.markdown(f"""<div class="kpi">
<div class="v" style="color:{color_c};">{s['cpc']} / {s['cpc_max']}</div>
<div class="l">Postes CPC ({pct_c}%)</div></div>""", unsafe_allow_html=True)

                            # Message qualitatif
                            total_pct = round((s['actif'] + s['passif'] + s['cpc'])
                                              / (s['actif_max'] + s['passif_max'] + s['cpc_max']) * 100)
                            if total_pct >= 60:
                                st.markdown(f'<div class="ok">✅ <strong>Bonne extraction</strong> — {total_pct}% des postes détectés. Vous pouvez générer l\'Excel.</div>',
                                            unsafe_allow_html=True)
                            elif total_pct >= 30:
                                st.markdown(f'<div class="warn">⚠️ <strong>Extraction partielle</strong> — {total_pct}% des postes détectés. Vérifiez les numéros de pages et réessayez.</div>',
                                            unsafe_allow_html=True)
                            else:
                                st.markdown(f'<div class="er">❌ <strong>Extraction insuffisante</strong> — {total_pct}% seulement. Vérifiez les pages indiquées.</div>',
                                            unsafe_allow_html=True)

                            # Stocker en session pour l'étape 5
                            st.session_state['parsed_rapport'] = parsed_r
                            st.session_state['pdf_path_r']     = pdf_path_r

                        except Exception as e:
                            logger.exception("Erreur analyse rapport")
                            st.markdown(
                                f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                                unsafe_allow_html=True)
                            import traceback
                            st.code(traceback.format_exc())

                # ── ÉTAPE 5 : Générer l'Excel ─────────────────────
                if 'parsed_rapport' in st.session_state:
                    st.markdown("---")
                    st.markdown("### 📥 Étape 5 — Générer l'Excel")

                    # Validation identification
                    parsed_r = st.session_state['parsed_rapport']
                    info_check = parsed_r['info']
                    errors_r = []
                    if not info_check.get('raison_sociale'):  errors_r.append("Raison sociale")
                    if not info_check.get('exercice_fin'):    errors_r.append("Date de bilan")
                    if not info_check.get('centre_affaires'): errors_r.append("Centre d'affaires")
                    if info_check.get('macro_secteur') == "— Sélectionner —":
                        errors_r.append("Macro-secteur")

                    if errors_r:
                        st.markdown(
                            f'<div class="warn">⚠️ Complétez d\'abord : <strong>{", ".join(errors_r)}</strong> (Étape 2)</div>',
                            unsafe_allow_html=True)
                    else:
                        if st.button("📥 Générer l'Excel", type="primary",
                                     use_container_width=True, key="btn_generer"):
                            try:
                                from core.excel_writer import write

                                output_path_r = pdf_path_r.replace(".pdf", "_rapport_out.xlsx")
                                with st.spinner("Génération de l'Excel..."):
                                    stats_r = write(parsed_r, output_path_r)

                                raison_r = info_check.get('raison_sociale', 'RAPPORT')
                                date_r   = info_check.get('exercice_fin', '').replace('/', '_')
                                fname_r  = (
                                    f"FiscalXL_{raison_r.replace(' ','_')[:20]}"
                                    f"_{date_r}_Rapport.xlsx"
                                )

                                with open(output_path_r, "rb") as f_dl:
                                    st.download_button(
                                        "📥 Télécharger l'Excel",
                                        data=f_dl,
                                        file_name=fname_r,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="dl_rapport"
                                    )
                                st.markdown(f"""<div class="ok">
✅ <strong>Excel prêt</strong> &nbsp;·&nbsp;
<strong>{raison_r[:30]}</strong> &nbsp;·&nbsp;
{info_check.get('exercice_fin','')} &nbsp;·&nbsp;
Actif {stats_r['actif']} lignes · Passif {stats_r['passif']} · CPC {stats_r['cpc']}
</div>""", unsafe_allow_html=True)

                                # Nettoyage
                                try:
                                    if os.path.exists(output_path_r):
                                        os.unlink(output_path_r)
                                except Exception:
                                    pass

                            except Exception as e:
                                logger.exception("Erreur génération Excel rapport")
                                st.markdown(
                                    f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                                    unsafe_allow_html=True)
                                import traceback
                                st.code(traceback.format_exc())


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 2 — GUIDE D'UTILISATION
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("## 📖 Guide d'utilisation")
    st.markdown("Ce guide explique comment utiliser FiscalXL Pro.")
    st.markdown("---")

    st.markdown("### Étape 1 — Choisir le format du PDF")
    st.markdown("""
Dans la **barre latérale gauche**, sélectionnez le format de votre PDF :

| Format | Pages | Utilisation |
|--------|-------|-------------|
| **AMMC** | 5 pages | Bilans déposés auprès de l'AMMC |
| **DGI** | 7 pages | États de synthèse DGI |
| **Rapport financier** | Pages libres | Rapports annuels (LabelVie, ONCF...) |
""")
    st.markdown("---")

    st.markdown("### Étape 2 — Mode Rapport financier")
    st.markdown("""
Pour le mode **Rapport financier** :
1. Uploadez le PDF du rapport
2. Renseignez **manuellement** les informations d'identification
3. Indiquez les **numéros de pages** contenant Actif, Passif, CPC
4. Cliquez sur **Analyser** pour vérifier l'extraction
5. Si le résultat est satisfaisant (≥ 60%), générez l'Excel
""")
    st.info("💡 Si l'extraction est insuffisante, vérifiez que les pages indiquées correspondent bien aux tableaux financiers et réessayez.")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 3 — DOCUMENTATION TECHNIQUE
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("## 🔧 Documentation technique")
    st.markdown("Documentation à destination de l'équipe Gestion des Risques.")
    st.markdown("---")

    st.markdown("### Architecture du projet")
    st.code("""
FiscalXL_Pro/
├── app.py                  ← Interface Streamlit
├── core/
│   ├── ammc_parser.py      ← Parser PDF format AMMC (5 pages)
│   ├── dgi_parser.py       ← Parser PDF format DGI (7 pages)
│   ├── rapport_parser.py   ← Parser rapports financiers libres (NOUVEAU)
│   ├── excel_writer.py     ← Génération Excel (commun tous formats)
│   └── synonyms.py         ← Dictionnaire des variantes de labels
├── utils/
│   └── logger.py           ← Journalisation
└── requirements.txt
    """, language="")

    st.markdown("---")

    st.markdown("### Formats PDF supportés")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="doc-section"><h4>📄 AMMC</h4>', unsafe_allow_html=True)
        st.markdown("5 pages fixes · Template MCN · `ammc_parser.py`")
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="doc-section"><h4>🏛️ DGI</h4>', unsafe_allow_html=True)
        st.markdown("7 pages fixes · Template MCN · `dgi_parser.py`")
        st.markdown('</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="doc-section"><h4>📑 Rapport</h4>', unsafe_allow_html=True)
        st.markdown("Pages libres · Pages indiquées manuellement · `rapport_parser.py`")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

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

    st.markdown("### Cas limites")
    st.markdown("""
| Situation | Comportement |
|-----------|-------------|
| PDF avec cellules fusionnées | Extraction X/Y automatique |
| Label non reconnu | Ignoré |
| Valeur manquante | Affiché comme `0` |
| CPC 6 colonnes DGI (Nature\|Label\|...) | Détection automatique `_is_cpc_6col()` |
| Extraction < 30% | Avertissement rouge — vérifier les pages |
""")
