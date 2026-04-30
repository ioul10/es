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
    Mode Rapport Financier — à intégrer dans app.py (bloc `else` du mode rapport)
    Remplace le bloc `else: # MODE RAPPORT FINANCIER` existant.
    """

    # ══════════════════════════════════════════════════════════════════
    # MODE RAPPORT FINANCIER
    # ══════════════════════════════════════════════════════════════════

    st.markdown("""<div class="rapport-box">
    <span class="step-badge-rapport">Rapport financier</span>
    <h3>Conversion d'un rapport financier libre</h3>
    <p>
    Uploadez votre rapport, indiquez les pages et les zones visuellement,
    puis générez l'Excel MCN standardisé.
    </p>
    </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── ÉTAPE 1 : Upload ─────────────────────────────────────────────
    st.markdown("### 📂 Étape 1 — Importer le rapport PDF")
    uploaded_r = st.file_uploader(
        "Rapport financier annuel (PDF)",
        type=["pdf"],
        key="upload_rapport"
    )

    if not uploaded_r:
        st.markdown("""<div style="text-align:center;padding:2.5rem;color:#888;
    border:2px dashed #70AD47;border-radius:12px;background:#f0f7f0;">
    <div style="font-size:3rem;">📑</div>
    <h3 style="color:#375623;">Glissez votre rapport PDF ici</h3>
    <p>Rapport financier annuel · Pages libres</p>
    </div>""", unsafe_allow_html=True)

    else:
        # Sauvegarder le PDF en session
        if 'pdf_bytes_r' not in st.session_state or st.session_state.get('pdf_name_r') != uploaded_r.name:
            st.session_state['pdf_bytes_r'] = uploaded_r.getbuffer().tobytes()
            st.session_state['pdf_name_r']  = uploaded_r.name
            # Reset parsing si nouveau fichier
            for k in ['parsed_rapport', 'pdf_path_r']:
                st.session_state.pop(k, None)

        import tempfile, os, io
        import pdfplumber
        from PIL import Image

        # Écrire le PDF dans un fichier temp permanent (pour pdfplumber)
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
            st.success(f"✅ PDF chargé — **{total_pages} page{'s' if total_pages > 1 else ''}** détectée{'s' if total_pages > 1 else ''}")
        except Exception as e:
            st.error(f"Erreur lecture PDF : {e}")
            total_pages = 0

        if total_pages > 0:
            st.markdown("---")

            # ── ÉTAPE 2 : Identification ──────────────────────────────
            st.markdown("### ✏️ Étape 2 — Informations d'identification")

            col_r1, col_r2 = st.columns(2)
            with col_r1:
                r_raison  = st.text_input("Raison sociale *", placeholder="Ex : LABELVIE SA", key="r_raison")
                r_if      = st.text_input("Identifiant fiscal", placeholder="Ex : 4510887", key="r_if")
                r_taxe    = st.text_input("Taxe professionnelle", placeholder="Ex : 25725940", key="r_taxe")
            with col_r2:
                r_date    = st.text_input("Date de bilan *", placeholder="Ex : 31/12/2025", key="r_date")
                r_centre  = st.text_input("Centre d'affaires *", placeholder="Ex : Casa Finance", key="r_centre")
                r_secteur = st.selectbox("Macro-secteur *",
                    options=["— Sélectionner —","Manufactures","Services","Commerce","BTP","Holding"],
                    key="r_secteur")

            st.markdown("---")

            # ── ÉTAPE 3 : Pages + Zones visuelles ────────────────────
            st.markdown("### 🗺️ Étape 3 — Localiser les tableaux")
            st.caption(
                "Pour chaque section, indiquez la page et la zone horizontale "
                "en glissant les curseurs. **0% = bord gauche · 100% = bord droit.** "
                "Si le tableau occupe toute la largeur, laissez 0%→100%."
            )

            # Sélecteur de page pour l'aperçu
            col_prev, col_info = st.columns([2, 1])
            with col_prev:
                preview_page = st.selectbox(
                    "📄 Aperçu de la page",
                    options=list(range(1, total_pages + 1)),
                    format_func=lambda x: f"Page {x}",
                    key="preview_page"
                )
            with col_info:
                w, h = page_dims[preview_page - 1]
                st.markdown(f"""<div style="background:#f0f7f0;border-radius:8px;
    padding:.8rem;margin-top:1.6rem;font-size:.85rem;color:#375623;">
    📐 <strong>{w:.0f} × {h:.0f} pt</strong><br>
    🔢 Page {preview_page} / {total_pages}
    </div>""", unsafe_allow_html=True)

            # Génération aperçu
            @st.cache_data(show_spinner=False)
            def get_page_preview(pdf_path: str, page_idx: int, resolution: int = 80) -> bytes:
                with pdfplumber.open(pdf_path) as pdf:
                    img = pdf.pages[page_idx].to_image(resolution=resolution)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    return buf.getvalue()

            with st.spinner("Génération de l'aperçu..."):
                preview_bytes = get_page_preview(pdf_path_r, preview_page - 1)

            # Afficher l'aperçu avec les zones colorées
            col_img, col_zones = st.columns([1, 1])

            with col_img:
                st.image(preview_bytes, caption=f"Page {preview_page}", use_column_width=True)

            with col_zones:
                st.markdown("#### 📊 Définir les zones")

                SECTIONS = [
                    ("actif",  "🟦 Bilan Actif",  "#2E75B6"),
                    ("passif", "🟩 Bilan Passif",  "#70AD47"),
                    ("cpc",    "🟧 CPC",           "#ED7D31"),
                ]

                zone_config = {}
                for key, label, color in SECTIONS:
                    st.markdown(f"**{label}**")
                    p_col, _ = st.columns([1, 1])
                    with p_col:
                        pages_val = st.text_input(
                            f"Pages",
                            value="1" if total_pages == 1 else "",
                            placeholder="Ex: 1 ou 1,2",
                            key=f"pages_{key}",
                            label_visibility="collapsed"
                        )
                        st.caption(f"Pages pour {label.split()[1]}")

                    zone_pct = st.slider(
                        f"Zone horizontale",
                        min_value=0, max_value=100,
                        value=(0, 100),
                        step=5,
                        key=f"zone_{key}",
                        label_visibility="collapsed",
                        help=f"Glisser pour délimiter la zone du tableau {label}"
                    )

                    # Afficher la zone choisie
                    pct_left, pct_right = zone_pct
                    bar_html = f"""
    <div style="background:#e0e0e0;border-radius:4px;height:12px;margin:2px 0 8px;">
      <div style="background:{color};border-radius:4px;height:12px;
        margin-left:{pct_left}%;width:{pct_right-pct_left}%;"></div>
    </div>
    <div style="font-size:.75rem;color:#666;margin-bottom:12px;">
      Zone : {pct_left}% → {pct_right}%
    </div>"""
                    st.markdown(bar_html, unsafe_allow_html=True)

                    zone_config[key] = {
                        'pages': pages_val,
                        'pct_left':  pct_left,
                        'pct_right': pct_right,
                    }

            # Overlay visuel des 3 zones sur l'aperçu
            st.markdown("#### 🎨 Aperçu des zones sélectionnées")

            # Recréer l'image avec overlay des zones
            @st.cache_data(show_spinner=False)
            def get_page_preview_hd(pdf_path: str, page_idx: int) -> tuple:
                with pdfplumber.open(pdf_path) as pdf:
                    page = pdf.pages[page_idx]
                    img = page.to_image(resolution=100)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    return buf.getvalue(), page.width, page.height

            img_bytes_hd, page_w, page_h = get_page_preview_hd(pdf_path_r, preview_page - 1)

            # Dessiner les zones avec PIL
            from PIL import Image, ImageDraw, ImageFont
            img_pil = Image.open(io.BytesIO(img_bytes_hd)).convert("RGBA")
            overlay = Image.new("RGBA", img_pil.size, (0, 0, 0, 0))
            draw    = ImageDraw.Draw(overlay)

            COLORS_RGBA = {
                'actif':  (46,  117, 182, 60),
                'passif': (112, 173,  71, 60),
                'cpc':    (237, 125,  49, 60),
            }
            COLORS_BORDER = {
                'actif':  (46,  117, 182, 200),
                'passif': (112, 173,  71, 200),
                'cpc':    (237, 125,  49, 200),
            }

            img_w, img_h = img_pil.size

            for key, cfg in zone_config.items():
                # Convertir % → pixels image
                x0_px = int(cfg['pct_left']  / 100 * img_w)
                x1_px = int(cfg['pct_right'] / 100 * img_w)
                if x1_px > x0_px:
                    draw.rectangle([x0_px, 0, x1_px, img_h],
                                   fill=COLORS_RGBA[key])
                    draw.rectangle([x0_px, 0, x1_px, img_h],
                                   outline=COLORS_BORDER[key], width=3)
                    # Label
                    label_map = {'actif': 'ACTIF', 'passif': 'PASSIF', 'cpc': 'CPC'}
                    draw.text((x0_px + 5, 10), label_map[key],
                              fill=COLORS_BORDER[key])

            composite = Image.alpha_composite(img_pil, overlay).convert("RGB")
            buf_out = io.BytesIO()
            composite.save(buf_out, format="PNG")

            col_ov1, col_ov2, col_ov3 = st.columns([1, 2, 1])
            with col_ov2:
                st.image(buf_out.getvalue(),
                         caption="Zones définies — Bleu=Actif · Vert=Passif · Orange=CPC",
                         use_column_width=True)

            # Légende
            st.markdown("""
    <div style="display:flex;gap:1.5rem;margin:.5rem 0;">
    <span style="color:#2E75B6;font-weight:bold;">🟦 Actif</span>
    <span style="color:#70AD47;font-weight:bold;">🟩 Passif</span>
    <span style="color:#ED7D31;font-weight:bold;">🟧 CPC</span>
    </div>""", unsafe_allow_html=True)

            st.markdown("---")

            # ── ÉTAPE 4 : Analyser ────────────────────────────────────
            st.markdown("### 🔍 Étape 4 — Analyser")

            pages_ok = all(zone_config[k]['pages'].strip() for k in ['actif','passif','cpc'])
            if not pages_ok:
                st.markdown('<div class="warn">⚠️ Renseignez les pages pour les 3 sections avant d\'analyser.</div>',
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

                    # Construire les paramètres de zones
                    zones = {
                        k: {
                            'pages':     zone_config[k]['pages'],
                            'pct_left':  zone_config[k]['pct_left'],
                            'pct_right': zone_config[k]['pct_right'],
                        }
                        for k in ['actif', 'passif', 'cpc']
                    }

                    with st.spinner("Analyse en cours..."):
                        parsed_r = parse_rapport(
                            pdf_path_r,
                            pages_actif=zones['actif']['pages'],
                            pages_passif=zones['passif']['pages'],
                            pages_cpc=zones['cpc']['pages'],
                            zone_actif=(zones['actif']['pct_left'],   zones['actif']['pct_right']),
                            zone_passif=(zones['passif']['pct_left'], zones['passif']['pct_right']),
                            zone_cpc=(zones['cpc']['pct_left'],       zones['cpc']['pct_right']),
                            info=info_r,
                        )

                    s = parsed_r['_stats']
                    st.markdown("#### Résultat de l'analyse")
                    ka, kp, kc = st.columns(3)

                    def kpi_card(col, label, found, total):
                        pct = round(found / total * 100)
                        color = "#375623" if pct >= 60 else ("#7F6000" if pct >= 30 else "#7B2C00")
                        col.markdown(f"""<div class="kpi">
    <div class="v" style="color:{color};">{found} / {total}</div>
    <div class="l">{label} ({pct}%)</div></div>""", unsafe_allow_html=True)

                    kpi_card(ka, "Actif",  s['actif'],  s['actif_max'])
                    kpi_card(kp, "Passif", s['passif'], s['passif_max'])
                    kpi_card(kc, "CPC",    s['cpc'],    s['cpc_max'])

                    total_pct = round(
                        (s['actif'] + s['passif'] + s['cpc']) /
                        (s['actif_max'] + s['passif_max'] + s['cpc_max']) * 100
                    )

                    if total_pct >= 60:
                        st.markdown(f'<div class="ok">✅ <strong>Bonne extraction</strong> — {total_pct}% des postes détectés.</div>',
                                    unsafe_allow_html=True)
                    elif total_pct >= 30:
                        st.markdown(f'<div class="warn">⚠️ <strong>Extraction partielle</strong> — {total_pct}%. Ajustez les zones et réessayez.</div>',
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="er">❌ <strong>Extraction insuffisante</strong> — {total_pct}%. Vérifiez les pages et zones.</div>',
                                    unsafe_allow_html=True)

                    st.session_state['parsed_rapport'] = parsed_r

                except Exception as e:
                    logger.exception("Erreur analyse rapport")
                    st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                                unsafe_allow_html=True)
                    import traceback
                    st.code(traceback.format_exc())

            # ── ÉTAPE 5 : Générer Excel ───────────────────────────────
            if 'parsed_rapport' in st.session_state:
                st.markdown("---")
                st.markdown("### 📥 Étape 5 — Générer l'Excel")

                parsed_r  = st.session_state['parsed_rapport']
                info_check = parsed_r['info']
                errors_r   = []
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
                            fname_r  = f"FiscalXL_{raison_r.replace(' ','_')[:20]}_{date_r}_Rapport.xlsx"

                            with open(output_path_r, "rb") as f_dl:
                                st.download_button(
                                    "📥 Télécharger l'Excel", data=f_dl,
                                    file_name=fname_r,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_rapport"
                                )
                            st.markdown(f"""<div class="ok">
    ✅ <strong>Excel prêt</strong> &nbsp;·&nbsp;
    <strong>{raison_r[:30]}</strong> &nbsp;·&nbsp;
    {info_check.get('exercice_fin','')} &nbsp;·&nbsp;
    Actif {stats_r['actif']} · Passif {stats_r['passif']} · CPC {stats_r['cpc']}
    </div>""", unsafe_allow_html=True)

                            try:
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
