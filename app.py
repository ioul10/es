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
    """
    Mode Rapport Financier v2 — avec détection automatique des tableaux
    Remplace le bloc `else: # MODE RAPPORT FINANCIER` dans app.py
    """

    # ══════════════════════════════════════════════════════════════════
    # MODE RAPPORT FINANCIER
    # ══════════════════════════════════════════════════════════════════

    import io, os, tempfile
    import pdfplumber
    from PIL import Image, ImageDraw

    st.markdown("""<div class="rapport-box">
    <span class="step-badge-rapport">Rapport financier</span>
    <h3>Conversion d'un rapport financier libre</h3>
    <p>Uploadez votre rapport — les tableaux sont détectés automatiquement.
    Ajustez les zones si nécessaire puis générez l'Excel MCN.</p>
    </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── ÉTAPE 1 : Upload ─────────────────────────────────────────────
    st.markdown("### 📂 Étape 1 — Importer le rapport PDF")

    uploaded_r = st.file_uploader(
        "Rapport financier annuel (PDF)",
        type=["pdf"], key="upload_rapport"
    )

    if not uploaded_r:
        st.markdown("""<div style="text-align:center;padding:2.5rem;color:#888;
    border:2px dashed #70AD47;border-radius:12px;background:#f0f7f0;">
    <div style="font-size:3rem;">📑</div>
    <h3 style="color:#375623;">Glissez votre rapport PDF ici</h3>
    <p>Rapport financier annuel · Détection automatique des tableaux</p>
    </div>""", unsafe_allow_html=True)

    else:
        # Sauvegarder PDF en session
        if (st.session_state.get('pdf_name_r') != uploaded_r.name):
            st.session_state['pdf_bytes_r'] = uploaded_r.getbuffer().tobytes()
            st.session_state['pdf_name_r']  = uploaded_r.name
            for k in ['parsed_rapport', 'detected_tables', 'zone_config']:
                st.session_state.pop(k, None)

        # Écrire dans fichier temp
        if 'pdf_path_r' not in st.session_state:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp.write(st.session_state['pdf_bytes_r'])
            tmp.close()
            st.session_state['pdf_path_r'] = tmp.name

        pdf_path_r = st.session_state['pdf_path_r']

        try:
            with pdfplumber.open(pdf_path_r) as _pdf:
                total_pages = len(_pdf.pages)
                page_dims   = [(p.width, p.height) for p in _pdf.pages]
            st.success(f"✅ PDF chargé — **{total_pages} page{'s' if total_pages>1 else ''}**")
        except Exception as e:
            st.error(f"Erreur lecture PDF : {e}")
            total_pages = 0

        if total_pages > 0:
            st.markdown("---")

            # ── ÉTAPE 2 : Identification ──────────────────────────────
            st.markdown("### ✏️ Étape 2 — Informations d'identification")
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                r_raison  = st.text_input("Raison sociale *",    placeholder="Ex : LABELVIE SA",    key="r_raison")
                r_if      = st.text_input("Identifiant fiscal",  placeholder="Ex : 4510887",         key="r_if")
                r_taxe    = st.text_input("Taxe professionnelle",placeholder="Ex : 25725940",        key="r_taxe")
            with col_r2:
                r_date    = st.text_input("Date de bilan *",     placeholder="Ex : 31/12/2025",      key="r_date")
                r_centre  = st.text_input("Centre d'affaires *", placeholder="Ex : Casa Finance",    key="r_centre")
                r_secteur = st.selectbox("Macro-secteur *",
                    options=["— Sélectionner —","Manufactures","Services","Commerce","BTP","Holding"],
                    key="r_secteur")

            st.markdown("---")

            # ── ÉTAPE 3 : Détection + Zones ──────────────────────────
            st.markdown("### 🗺️ Étape 3 — Localiser les tableaux")

            # Bouton de détection automatique
            col_det1, col_det2 = st.columns([1, 2])
            with col_det1:
                detect_clicked = st.button(
                    "🔍 Détecter automatiquement",
                    type="primary",
                    use_container_width=True,
                    key="btn_detect"
                )
            with col_det2:
                st.caption(
                    "Cliquez pour détecter les tableaux automatiquement. "
                    "Les zones seront pré-remplies — ajustez si nécessaire."
                )

            if detect_clicked:
                with st.spinner("Analyse du PDF en cours..."):
                    try:
                        from core.table_detector import detect_tables, summarize
                        tables   = detect_tables(pdf_path_r)
                        summary  = summarize(tables)
                        st.session_state['detected_tables'] = tables
                        st.session_state['detected_summary'] = summary

                        # Construire zone_config depuis le résumé
                        zc = {}
                        defaults = {'actif': (0,50), 'passif': (50,100), 'cpc': (0,100)}
                        pages_default = {'actif': '1', 'passif': '1' if total_pages==1 else '2',
                                         'cpc': '2' if total_pages >= 2 else '1'}
                        for sec in ['actif','passif','cpc']:
                            if sec in summary:
                                t = summary[sec]
                                zc[sec] = {
                                    'pages':    str(t['page']),
                                    'pct_left': t['pct_x0'],
                                    'pct_right':t['pct_x1'],
                                    'detected': True,
                                }
                            else:
                                zc[sec] = {
                                    'pages':    pages_default[sec],
                                    'pct_left': defaults[sec][0],
                                    'pct_right':defaults[sec][1],
                                    'detected': False,
                                }
                        st.session_state['zone_config'] = zc

                        # Afficher résultat détection
                        n_found = len([s for s in ['actif','passif','cpc'] if s in summary])
                        if n_found == 3:
                            st.markdown(f'<div class="ok">✅ <strong>3/3 tableaux détectés</strong> — Actif, Passif et CPC localisés automatiquement.</div>',
                                        unsafe_allow_html=True)
                        elif n_found > 0:
                            found = [s for s in ['actif','passif','cpc'] if s in summary]
                            missing = [s for s in ['actif','passif','cpc'] if s not in summary]
                            st.markdown(
                                f'<div class="warn">⚠️ <strong>{n_found}/3 détectés</strong> — '
                                f'Trouvés : {", ".join(found)}. '
                                f'Non trouvés : {", ".join(missing)} — ajustez manuellement.</div>',
                                unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="er">❌ Aucun tableau détecté automatiquement. Renseignez les zones manuellement.</div>',
                                        unsafe_allow_html=True)
                    except Exception as e:
                        st.error(f"Erreur détection : {e}")
                        import traceback; st.code(traceback.format_exc())

            # Initialiser zone_config si pas encore fait
            if 'zone_config' not in st.session_state:
                st.session_state['zone_config'] = {
                    'actif':  {'pages': '1',                                    'pct_left': 0,  'pct_right': 100, 'detected': False},
                    'passif': {'pages': '1' if total_pages==1 else '2',         'pct_left': 0,  'pct_right': 100, 'detected': False},
                    'cpc':    {'pages': '2' if total_pages >= 2 else '1',       'pct_left': 0,  'pct_right': 100, 'detected': False},
                }

            zc = st.session_state['zone_config']

            # ── Aperçu + sliders ──────────────────────────────────────
            st.markdown("#### Ajuster les zones")
            st.caption("Les curseurs définissent la zone horizontale de chaque tableau (0% = bord gauche, 100% = bord droit)")

            # Sélecteur page aperçu
            preview_page = st.selectbox(
                "📄 Page à visualiser",
                options=list(range(1, total_pages + 1)),
                format_func=lambda x: f"Page {x}",
                key="preview_page"
            )

            # Colonnes : sliders à gauche, aperçu à droite
            col_sliders, col_preview = st.columns([1, 1])

            SECTION_META = {
                'actif':  ('🟦 Bilan Actif',  '#2E75B6', (46, 117, 182)),
                'passif': ('🟩 Bilan Passif', '#70AD47', (112, 173, 71)),
                'cpc':    ('🟧 CPC',          '#ED7D31', (237, 125, 49)),
            }

            new_zc = {}
            with col_sliders:
                for sec in ['actif', 'passif', 'cpc']:
                    label, color_hex, color_rgb = SECTION_META[sec]
                    cfg = zc[sec]

                    detected_badge = (
                        ' <span style="background:#E2EFDA;color:#375623;'
                        'border-radius:4px;padding:1px 6px;font-size:.72rem;">✅ auto</span>'
                        if cfg.get('detected') else ''
                    )
                    st.markdown(
                        f'<div style="font-weight:bold;color:{color_hex};margin-bottom:4px;">'
                        f'{label}{detected_badge}</div>',
                        unsafe_allow_html=True
                    )

                    p_col, z_col = st.columns([1, 2])
                    with p_col:
                        pages_val = st.text_input(
                            "Pages", value=cfg['pages'],
                            placeholder="ex: 1 ou 1,2",
                            key=f"pages_{sec}",
                            label_visibility="collapsed"
                        )
                        st.caption("Pages")

                    with z_col:
                        zone_pct = st.slider(
                            "Zone", min_value=0, max_value=100,
                            value=(cfg['pct_left'], cfg['pct_right']),
                            step=5, key=f"zone_{sec}",
                            label_visibility="collapsed"
                        )

                    # Barre colorée
                    pct_l, pct_r = zone_pct
                    st.markdown(f"""
    <div style="background:#e8e8e8;border-radius:3px;height:8px;margin:0 0 12px;">
      <div style="background:{color_hex};border-radius:3px;height:8px;
        margin-left:{pct_l}%;width:{max(pct_r-pct_l,1)}%;"></div>
    </div>""", unsafe_allow_html=True)

                    new_zc[sec] = {
                        'pages':     pages_val,
                        'pct_left':  pct_l,
                        'pct_right': pct_r,
                        'detected':  cfg.get('detected', False),
                    }

            # Sauvegarder les nouvelles valeurs
            st.session_state['zone_config'] = new_zc

            # ── Aperçu avec overlay ───────────────────────────────────
            @st.cache_data(show_spinner=False)
            def get_preview(pdf_path: str, page_idx: int) -> bytes:
                with pdfplumber.open(pdf_path) as pdf:
                    img = pdf.pages[page_idx].to_image(resolution=90)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    return buf.getvalue()

            with col_preview:
                prev_bytes = get_preview(pdf_path_r, preview_page - 1)
                img_pil    = Image.open(io.BytesIO(prev_bytes)).convert("RGBA")
                overlay    = Image.new("RGBA", img_pil.size, (0,0,0,0))
                draw       = ImageDraw.Draw(overlay)
                iW, iH    = img_pil.size

                for sec in ['actif', 'passif', 'cpc']:
                    cfg = new_zc[sec]
                    _, _, rgb = SECTION_META[sec]

                    # Vérifier si cette section est sur la page affichée
                    try:
                        from core.rapport_parser import _parse_pages_input
                        pages_list = _parse_pages_input(cfg['pages'])
                    except Exception:
                        import re
                        pages_list = [int(p)-1 for p in re.split(r'[,;]', cfg['pages'])
                                      if p.strip().isdigit()]

                    if (preview_page - 1) not in pages_list:
                        continue

                    x0 = int(cfg['pct_left']  / 100 * iW)
                    x1 = int(cfg['pct_right'] / 100 * iW)
                    if x1 <= x0: continue

                    draw.rectangle([x0, 0, x1, iH],
                                   fill=(*rgb, 50),
                                   outline=(*rgb, 200))
                    draw.rectangle([x0, 0, x1, 22],
                                   fill=(*rgb, 180))
                    lbl_map = {'actif':'ACTIF','passif':'PASSIF','cpc':'CPC'}
                    draw.text((x0+4, 4), lbl_map[sec], fill=(255,255,255,255))

                composite = Image.alpha_composite(img_pil, overlay).convert("RGB")
                buf_out   = io.BytesIO()
                composite.save(buf_out, format="PNG")
                st.image(buf_out.getvalue(),
                         caption=f"Page {preview_page} — zones colorées",
                         use_column_width=True)

            st.markdown("---")

            # ── ÉTAPE 4 : Analyser ────────────────────────────────────
            st.markdown("### 🔍 Étape 4 — Analyser")

            pages_ok = all(new_zc[k]['pages'].strip() for k in ['actif','passif','cpc'])

            if not pages_ok:
                st.markdown('<div class="warn">⚠️ Renseignez les pages pour les 3 sections.</div>',
                            unsafe_allow_html=True)

            if st.button("🔍 Analyser le rapport", type="primary",
                         use_container_width=True, key="btn_analyser",
                         disabled=not pages_ok):
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
                            pages_actif  = new_zc['actif']['pages'],
                            pages_passif = new_zc['passif']['pages'],
                            pages_cpc    = new_zc['cpc']['pages'],
                            zone_actif   = (new_zc['actif']['pct_left'],   new_zc['actif']['pct_right']),
                            zone_passif  = (new_zc['passif']['pct_left'],  new_zc['passif']['pct_right']),
                            zone_cpc     = (new_zc['cpc']['pct_left'],     new_zc['cpc']['pct_right']),
                            info=info_r,
                        )

                    s = parsed_r['_stats']
                    st.markdown("#### Résultat")
                    ka, kp, kc = st.columns(3)

                    def kpi(col, label, found, total):
                        pct   = round(found/total*100)
                        color = "#375623" if pct>=60 else ("#7F6000" if pct>=30 else "#7B2C00")
                        col.markdown(f"""<div class="kpi">
    <div class="v" style="color:{color};">{found}/{total}</div>
    <div class="l">{label} ({pct}%)</div></div>""", unsafe_allow_html=True)

                    kpi(ka, "Actif",  s['actif'],  s['actif_max'])
                    kpi(kp, "Passif", s['passif'], s['passif_max'])
                    kpi(kc, "CPC",    s['cpc'],    s['cpc_max'])

                    total_pct = round(
                        (s['actif']+s['passif']+s['cpc']) /
                        (s['actif_max']+s['passif_max']+s['cpc_max']) * 100
                    )

                    if total_pct >= 60:
                        st.markdown(f'<div class="ok">✅ <strong>Bonne extraction</strong> — {total_pct}%.</div>',
                                    unsafe_allow_html=True)
                    elif total_pct >= 30:
                        st.markdown(f'<div class="warn">⚠️ <strong>Extraction partielle</strong> — {total_pct}%. Ajustez les zones.</div>',
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="er">❌ <strong>Insuffisant</strong> — {total_pct}%. Vérifiez les pages et zones.</div>',
                                    unsafe_allow_html=True)

                    st.session_state['parsed_rapport'] = parsed_r

                except Exception as e:
                    logger.exception("Erreur analyse rapport")
                    st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                                unsafe_allow_html=True)
                    import traceback; st.code(traceback.format_exc())

            # ── ÉTAPE 5 : Générer Excel ───────────────────────────────
            if 'parsed_rapport' in st.session_state:
                st.markdown("---")
                st.markdown("### 📥 Étape 5 — Générer l'Excel")

                parsed_r   = st.session_state['parsed_rapport']
                info_check = parsed_r['info']
                errors_r   = []
                if not info_check.get('raison_sociale'):  errors_r.append("Raison sociale")
                if not info_check.get('exercice_fin'):    errors_r.append("Date de bilan")
                if not info_check.get('centre_affaires'): errors_r.append("Centre d'affaires")
                if info_check.get('macro_secteur') == "— Sélectionner —":
                    errors_r.append("Macro-secteur")

                if errors_r:
                    st.markdown(
                        f'<div class="warn">⚠️ Complétez : <strong>{", ".join(errors_r)}</strong></div>',
                        unsafe_allow_html=True)
                else:
                    if st.button("📥 Générer l'Excel", type="primary",
                                 use_container_width=True, key="btn_generer"):
                        try:
                            from core.excel_writer import write
                            output_path_r = pdf_path_r.replace(".pdf", "_out.xlsx")
                            with st.spinner("Génération..."):
                                stats_r = write(parsed_r, output_path_r)

                            raison_r = info_check.get('raison_sociale','RAPPORT')
                            date_r   = info_check.get('exercice_fin','').replace('/','_')
                            fname_r  = f"FiscalXL_{raison_r.replace(' ','_')[:20]}_{date_r}_Rapport.xlsx"

                            with open(output_path_r, "rb") as f_dl:
                                st.download_button(
                                    "📥 Télécharger l'Excel", data=f_dl,
                                    file_name=fname_r,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_rapport"
                                )
                            st.markdown(f"""<div class="ok">
    ✅ <strong>Excel prêt</strong> · <strong>{raison_r[:25]}</strong> ·
    {info_check.get('exercice_fin','')} ·
    Actif {stats_r['actif']} · Passif {stats_r['passif']} · CPC {stats_r['cpc']}
    </div>""", unsafe_allow_html=True)
                            try: os.unlink(output_path_r)
                            except Exception: pass

                        except Exception as e:
                            logger.exception("Erreur génération Excel")
                            st.markdown(f'<div class="er">❌ <code>{e}</code></div>',
                                        unsafe_allow_html=True)
                            import traceback; st.code(traceback.format_exc())

# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 2
