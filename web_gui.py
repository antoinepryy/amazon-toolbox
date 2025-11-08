# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# Configure page
st.set_page_config(
    page_title="Convertisseur CSV vers Excel",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better design
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #2c3e50;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 2rem;
    }
    
    .upload-section {
        background-color: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        border: 2px dashed #dee2e6;
        margin: 1rem 0;
        text-align: center;
    }
    
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #28a745;
        margin: 1rem 0;
    }
    
    .info-message {
        background-color: #d1ecf1;
        color: #0c5460;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #17a2b8;
        margin: 1rem 0;
    }
    
    .stDownloadButton > button {
        background-color: #28a745;
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.5rem 2rem;
        border-radius: 5px;
        font-size: 1.1rem;
    }
</style>
""", unsafe_allow_html=True)

def convert_csv_to_excel(csv_file):
    """
    Convertit un fichier CSV au format Excel avec gestion appropriée du format CSV Amazon
    """
    try:
        # Lire le contenu du fichier téléchargé
        csv_content = csv_file.read().decode('utf-8-sig')
        lines = csv_content.strip().split('\n')
        
        # Corriger le formatage CSV malformé
        fixed_lines = []
        for i, line in enumerate(lines):
            line = line.strip()
            if i == 0:  # Ligne d'en-tête
                fixed_lines.append(line)
            else:  # Lignes de données - gérer les champs entre guillemets
                if line.startswith('"') and line.endswith('"'):
                    # Supprimer les guillemets extérieurs
                    line = line[1:-1]
                    # Corriger les doubles guillemets intérieurs
                    line = line.replace('""', '"')
                fixed_lines.append(line)
        
        # Créer le DataFrame
        csv_string = '\n'.join(fixed_lines)
        df = pd.read_csv(io.StringIO(csv_string), sep=',')
        
        # Créer le fichier Excel en mémoire
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Données Amazon', index=False)
            
            # Ajuster automatiquement la largeur des colonnes
            worksheet = writer.sheets['Données Amazon']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        return output.getvalue(), df.shape
        
    except Exception as e:
        st.error(f"Erreur lors de la conversion du fichier : {str(e)}")
        return None, None

def main():
    # En-tête
    st.markdown('<h1 class="main-header">📊 Convertisseur CSV vers Excel</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-message">
        <strong>🎯 Objectif :</strong> Convertir les rapports CSV Amazon au format Excel avec formatage approprié des colonnes
    </div>
    """, unsafe_allow_html=True)
    
    # Téléchargeur de fichiers
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown("### 📁 Téléchargez votre fichier CSV")
    
    uploaded_file = st.file_uploader(
        "Choisissez un fichier CSV",
        type=['csv'],
        help="Sélectionnez votre rapport CSV Amazon ou tout fichier CSV à convertir au format Excel"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        # Afficher les détails du fichier
        st.markdown("### 📄 Informations du fichier")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Nom du fichier", uploaded_file.name)
        with col2:
            st.metric("Taille du fichier", f"{uploaded_file.size / 1024:.1f} KB")
        with col3:
            st.metric("Type de fichier", uploaded_file.type)
        
        # Aperçu des données CSV
        if st.checkbox("👀 Aperçu des données CSV", help="Afficher les 5 premières lignes de votre fichier CSV"):
            try:
                # Lire pour l'aperçu (réinitialiser le pointeur de fichier)
                uploaded_file.seek(0)
                preview_df = pd.read_csv(uploaded_file)
                st.dataframe(preview_df.head(), use_container_width=True)
                st.caption(f"Affichage des 5 premières lignes sur {len(preview_df)} lignes totales et {len(preview_df.columns)} colonnes")
            except Exception as e:
                st.warning(f"Impossible d'afficher l'aperçu du fichier : {str(e)}")
        
        # Bouton de conversion
        st.markdown("### 🔄 Convertir en Excel")
        
        if st.button("Convertir en Excel", type="primary", use_container_width=True):
            with st.spinner("Conversion de votre fichier CSV au format Excel en cours..."):
                # Réinitialiser le pointeur de fichier pour la conversion
                uploaded_file.seek(0)
                excel_data, shape = convert_csv_to_excel(uploaded_file)
                
                if excel_data:
                    # Stocker dans l'état de session pour le téléchargement
                    st.session_state.excel_data = excel_data
                    st.session_state.original_filename = uploaded_file.name
                    st.session_state.conversion_time = datetime.now()
                    st.session_state.data_shape = shape
                    
                    st.markdown(f"""
                    <div class="success-message">
                        <strong>✅ Conversion réussie !</strong><br>
                        Votre fichier CSV a été converti au format Excel.<br>
                        📊 Données : {shape[0]} lignes × {shape[1]} colonnes<br>
                        ⏰ Converti le : {st.session_state.conversion_time.strftime("%d/%m/%Y à %H:%M:%S")}
                    </div>
                    """, unsafe_allow_html=True)
    
    # Section de téléchargement
    if 'excel_data' in st.session_state:
        st.markdown("### 💾 Télécharger le fichier Excel")
        
        # Créer le nom de fichier
        original_name = os.path.splitext(st.session_state.original_filename)[0]
        excel_filename = f"{original_name}_converti.xlsx"
        
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=st.session_state.excel_data,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Informations supplémentaires
        st.info(f"💡 **Astuce :** Le fichier Excel sera sauvegardé sous le nom '{excel_filename}' dans votre dossier de téléchargements.")
    
    # Pied de page
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #6c757d; margin-top: 2rem;">
        <small>
            🛠️ <strong>Convertisseur CSV vers Excel</strong> | 
            Convertit les rapports CSV Amazon et autres fichiers CSV au format Excel avec formatage approprié
        </small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()