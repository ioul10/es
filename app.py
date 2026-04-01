"""Application Streamlit"""

import streamlit as st
import pandas as pd
from io import BytesIO
import time
from datetime import datetime

from core.extractor import FiscalPDFExtractor
from core.models import DocumentType
from config.settings import ExtractionConfig

# Configuration de la page
st.set_page_config(
    page_title="Convertisseur PDF Fiscal Marocain",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-message {
        padding: 1rem;
        background-color: #d4edda;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .warning-message {
        padding: 1rem;
        background-color: #fff3cd;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # En-tête
    st.markdown('<p class="main-header">📊 Convertisseur PDF Fiscal Marocain</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        doc_type = st.radio(
            "Type de document",
            ["Auto-détection", "AMMC", "DGI"],
            help="Sélectionnez le type ou laissez l'auto-détection"
        )
        
        st.markdown("---")
        
        st.subheader("📁 Options d'export")
        
        include_stats = st.checkbox("Inclure les statistiques", value=True)
        include_warnings = st.checkbox("Inclure les avertissements", value=True)
        
        st.markdown("---")
        
        st.subheader("🎯 Paramètres avancés")
        
        confidence_threshold = st.slider(
            "Seuil de confiance",
            min_value=0.0,
            max_value=1.0,
            value=0.5,
            step=0.05,
            help="En dessous de ce seuil, un avertissement sera affiché"
        )
        
        st.markdown("---")
        
        st.info("""
        **Formats supportés:**
        - ✅ AMMC (Modèle Comptable Normal)
        - ✅ DGI (Déclaration fiscale IS)
        
        **Données extraites:**
        - 📋 Identification
        - 📊 Bilan Actif
        - 📈 Bilan Passif
        - 📉 Compte de Produits et Charges
        """)
    
    # Zone principale
    uploaded_file = st.file_uploader(
        "📂 Déposez votre fichier PDF fiscal",
        type=["pdf"],
        help="Glissez-déposez ou cliquez pour sélectionner"
    )
    
    if uploaded_file:
        # Affichage des infos du fichier
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Nom du fichier", uploaded_file.name[:50] + "..." if len(uploaded_file.name) > 50 else uploaded_file.name)
        with col2:
            st.metric("📏 Taille", f"{uploaded_file.size / 1024:.1f} KB")
        with col3:
            st.metric("📅 Date", datetime.now().strftime("%d/%m/%Y %H:%M"))
        
        st.markdown("---")
        
        # Bouton de traitement
        if st.button("🚀 Démarrer l'extraction", type="primary", use_container_width=True):
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Déterminer le type de document
                selected_type = None
                if doc_type == "AMMC":
                    selected_type = DocumentType.AMMC
                elif doc_type == "DGI":
                    selected_type = DocumentType.DGI
                
                # Configurer l'extracteur
                config = ExtractionConfig()
                extractor = FiscalPDFExtractor(uploaded_file, selected_type, config)
                
                # Callback de progression
                def update_progress(current, total):
                    progress = current / total
                    progress_bar.progress(progress)
                    status_text.text(f"📊 Traitement des pages... {current}/{total}")
                
                # Extraction
                start_time = time.time()
                result = extractor.extract_all(progress_callback=update_progress)
                elapsed_time = time.time() - start_time
                
                # Nettoyer la progress bar
                progress_bar.empty()
                status_text.empty()
                
                # Vérifier le score de confiance
                if result.confidence_score < confidence_threshold:
                    st.warning(f"⚠️ Score de confiance bas: {result.confidence_score:.1%}")
                
                # Afficher le succès
                st.markdown(
                    f'<div class="success-message">✅ Extraction terminée en {elapsed_time:.1f} secondes !</div>',
                    unsafe_allow_html=True
                )
                
                # Métriques
                st.subheader("📊 Résumé de l'extraction")
                
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("🏢 Type", result.document_type.value)
                with col2:
                    st.metric("📄 Pages", result.pages_processed)
                with col3:
                    st.metric("📋 Tableaux", result.tables_found)
                with col4:
                    st.metric("📈 Lignes Actif", len(result.bilan_actif))
                with col5:
                    st.metric("🎯 Confiance", f"{result.confidence_score:.1%}")
                
                # Affichage des données
                st.subheader("📋 Aperçu des données")
                
                tabs = st.tabs(["🏢 Identification", "📊 Bilan Actif", "📈 Bilan Passif", "📉 CPC", "⚠️ Avertissements"])
                
                with tabs[0]:
                    if result.identification:
                        ident_data = result.identification.to_dict()
                        ident_df = pd.DataFrame([ident_data]).T
                        ident_df.columns = ["Valeur"]
                        st.dataframe(ident_df, use_container_width=True)
                    else:
                        st.warning("Aucune donnée d'identification")
                
                with tabs[1]:
                    if result.bilan_actif:
                        df = pd.DataFrame([line.to_dict() for line in result.bilan_actif])
                        st.dataframe(df, use_container_width=True, height=400)
                        
                        # Statistiques
                        col1, col2 = st.columns(2)
                        with col1:
                            total_brut = sum([l.brut for l in result.bilan_actif if l.brut])
                            st.metric("Total BRUT", f"{total_brut:,.2f} DH" if total_brut else "N/A")
                        with col2:
                            total_net = sum([l.net_n for l in result.bilan_actif if l.net_n])
                            st.metric("Total NET N", f"{total_net:,.2f} DH" if total_net else "N/A")
                    else:
                        st.warning("Aucune donnée bilan actif")
                
                with tabs[2]:
                    if result.bilan_passif:
                        df = pd.DataFrame([line.to_dict() for line in result.bilan_passif])
                        st.dataframe(df, use_container_width=True, height=400)
                    else:
                        st.warning("Aucune donnée bilan passif")
                
                with tabs[3]:
                    if result.cpc:
                        df = pd.DataFrame([line.to_dict() for line in result.cpc])
                        st.dataframe(df, use_container_width=True, height=400)
                        
                        # Résultat net
                        resultat_lines = [l for l in result.cpc if "RESULTAT NET" in l.designation.upper()]
                        if resultat_lines:
                            resultat = resultat_lines[0].total_n
                            st.metric("Résultat net de l'exercice", f"{resultat:,.2f} DH" if resultat else "N/A")
                    else:
                        st.warning("Aucune donnée CPC")
                
                with tabs[4]:
                    if result.warnings:
                        for warning in result.warnings:
                            st.warning(warning)
                    else:
                        st.success("✅ Aucun avertissement")
                
                if result.errors:
                    with st.expander("❌ Détails des erreurs"):
                        for error in result.errors:
                            st.error(error)
                
                # Export Excel
                st.markdown("---")
                
                excel_data = result.to_excel_data()
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet_name, data in excel_data.items():
                        if data:
                            df = pd.DataFrame(data)
                            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
                output.seek(0)
                
                # Nom du fichier
                filename = f"{result.identification.raison_sociale or 'document'}_{result.document_type.value}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                filename = filename.replace(" ", "_").replace("/", "-")
                
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ Erreur lors du traitement: {str(e)}")
                st.exception(e)
                progress_bar.empty()
                status_text.empty()

if __name__ == "__main__":
    main()
