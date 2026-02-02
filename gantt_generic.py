"""
Gantt Chart Generator - Application Streamlit
Génère des diagrammes de Gantt interactifs à partir de fichiers Excel
Colonnes requises: catégorie, tâche, début, fin
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import base64

# Configuration de la page
st.set_page_config(
    page_title="Gantt Generic",
    page_icon="📊",
    layout="wide"
)

# Styles CSS personnalisés
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    .stDownloadButton > button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)


def parse_date(date_val):
    """Parse une date depuis différents formats possibles."""
    if pd.isna(date_val):
        return None
    
    if isinstance(date_val, datetime):
        return date_val
    
    if isinstance(date_val, pd.Timestamp):
        return date_val.to_pydatetime()
    
    if isinstance(date_val, str):
        date_str = date_val.strip()
        formats = [
            "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y",
            "%Y/%m/%d", "%d.%m.%Y", "%Y.%m.%d"
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
    
    return None


def find_column(df, possible_names):
    """Trouve une colonne parmi plusieurs noms possibles."""
    df_cols_lower = {col.lower().strip(): col for col in df.columns}
    for name in possible_names:
        if name.lower() in df_cols_lower:
            return df_cols_lower[name.lower()]
    return None


def load_and_validate_data(uploaded_file):
    """Charge et valide les données du fichier Excel."""
    try:
        # Lecture du fichier Excel
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Recherche des colonnes requises
        col_mapping = {
            'categorie': find_column(df, ['catégorie', 'categorie', 'category', 'cat', 'groupe', 'group']),
            'tache': find_column(df, ['tâche', 'tache', 'task', 'nom', 'name', 'activité', 'activite']),
            'debut': find_column(df, ['début', 'debut', 'start', 'date_debut', 'date début', 'start_date']),
            'fin': find_column(df, ['fin', 'end', 'date_fin', 'date fin', 'end_date', 'échéance'])
        }
        
        # Vérification des colonnes manquantes
        missing = [k for k, v in col_mapping.items() if v is None]
        if missing:
            return None, f"Colonnes manquantes: {', '.join(missing)}. Colonnes disponibles: {', '.join(df.columns)}"
        
        # Création du DataFrame normalisé
        data = pd.DataFrame({
            'categorie': df[col_mapping['categorie']].astype(str).str.strip(),
            'tache': df[col_mapping['tache']].astype(str).str.strip(),
            'debut': df[col_mapping['debut']].apply(parse_date),
            'fin': df[col_mapping['fin']].apply(parse_date)
        })
        
        # Filtrage des lignes invalides
        data = data.dropna(subset=['debut', 'fin'])
        data = data[data['tache'].notna() & (data['tache'] != '') & (data['tache'] != 'nan')]
        
        # Calcul de la durée
        data['duree_jours'] = (data['fin'] - data['debut']).dt.days + 1
        
        # Tri par catégorie et date de début
        data = data.sort_values(['categorie', 'debut']).reset_index(drop=True)
        
        if len(data) == 0:
            return None, "Aucune donnée valide trouvée après filtrage."
        
        return data, None
        
    except Exception as e:
        return None, f"Erreur de lecture: {str(e)}"


def generate_color_palette(n):
    """Génère une palette de couleurs distinctes."""
    colors = [
        "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
        "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
        "#6b6ecf", "#b5cf6b", "#d6616b", "#ce6dbd", "#de9ed6"
    ]
    return [colors[i % len(colors)] for i in range(n)]


def generate_gantt_svg(data, title="Diagramme de Gantt"):
    """Génère un diagramme de Gantt en SVG."""
    if len(data) == 0:
        return "<svg><text>Aucune donnée</text></svg>"
    
    # Dimensions
    margin_left = 280
    margin_right = 40
    margin_top = 80
    margin_bottom = 60
    bar_height = 28
    bar_spacing = 8
    row_height = bar_height + bar_spacing
    
    # Calcul des dimensions totales
    n_tasks = len(data)
    chart_height = n_tasks * row_height
    total_height = margin_top + chart_height + margin_bottom
    total_width = 1200
    chart_width = total_width - margin_left - margin_right
    
    # Plage de dates
    min_date = data['debut'].min()
    max_date = data['fin'].max()
    date_range = (max_date - min_date).days
    if date_range <= 0:
        date_range = 1
    
    # Couleurs par catégorie
    categories = data['categorie'].unique()
    colors = generate_color_palette(len(categories))
    color_map = dict(zip(categories, colors))
    
    # Construction du SVG
    svg_parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total_width} {total_height}" '
        f'width="{total_width}" height="{total_height}">'
    ]
    
    # Styles
    svg_parts.append("""
    <defs>
        <style>
            .title { font: bold 20px sans-serif; fill: #1f4e79; }
            .axis-label { font: 11px sans-serif; fill: #333; }
            .task-label { font: 12px sans-serif; fill: #333; }
            .category-label { font: bold 11px sans-serif; fill: #666; }
            .grid-line { stroke: #e0e0e0; stroke-width: 1; }
            .bar { rx: 4; ry: 4; }
            .bar-text { font: 10px sans-serif; fill: white; }
            .legend-text { font: 11px sans-serif; fill: #333; }
        </style>
    </defs>
    """)
    
    # Fond
    svg_parts.append(f'<rect width="{total_width}" height="{total_height}" fill="white"/>')
    
    # Titre
    svg_parts.append(f'<text x="{total_width/2}" y="35" text-anchor="middle" class="title">{title}</text>')
    
    # Sous-titre avec période
    period_text = f"{min_date.strftime('%d/%m/%Y')} - {max_date.strftime('%d/%m/%Y')}"
    svg_parts.append(f'<text x="{total_width/2}" y="55" text-anchor="middle" class="axis-label">{period_text}</text>')
    
    # Grille verticale (dates)
    n_grid_lines = min(12, max(4, date_range // 30))
    for i in range(n_grid_lines + 1):
        x = margin_left + (i / n_grid_lines) * chart_width
        svg_parts.append(f'<line x1="{x}" y1="{margin_top}" x2="{x}" y2="{margin_top + chart_height}" class="grid-line"/>')
        
        # Labels de date
        grid_date = min_date + timedelta(days=int(i * date_range / n_grid_lines))
        date_label = grid_date.strftime('%d/%m/%y')
        svg_parts.append(f'<text x="{x}" y="{margin_top + chart_height + 20}" text-anchor="middle" class="axis-label">{date_label}</text>')
    
    # Grille horizontale et barres
    current_category = None
    for idx, row in data.iterrows():
        y = margin_top + idx * row_height
        
        # Ligne de grille horizontale
        svg_parts.append(f'<line x1="{margin_left}" y1="{y + row_height}" x2="{total_width - margin_right}" y2="{y + row_height}" class="grid-line"/>')
        
        # Indicateur de catégorie (changement)
        if row['categorie'] != current_category:
            current_category = row['categorie']
            cat_display = current_category[:25] + "..." if len(current_category) > 25 else current_category
            svg_parts.append(f'<text x="5" y="{y + bar_height/2 + 4}" class="category-label">{cat_display}</text>')
        
        # Label de tâche
        task_display = row['tache'][:30] + "..." if len(row['tache']) > 30 else row['tache']
        svg_parts.append(f'<text x="{margin_left - 10}" y="{y + bar_height/2 + 4}" text-anchor="end" class="task-label">{task_display}</text>')
        
        # Calcul position de la barre
        start_offset = (row['debut'] - min_date).days
        duration = (row['fin'] - row['debut']).days + 1
        
        bar_x = margin_left + (start_offset / date_range) * chart_width
        bar_width = max(4, (duration / date_range) * chart_width)
        
        color = color_map[row['categorie']]
        
        # Barre
        svg_parts.append(f'<rect x="{bar_x}" y="{y + 2}" width="{bar_width}" height="{bar_height - 4}" fill="{color}" class="bar">')
        svg_parts.append(f'<title>{row["tache"]} | {row["categorie"]} | {row["debut"].strftime("%d/%m/%Y")} → {row["fin"].strftime("%d/%m/%Y")} ({row["duree_jours"]}j)</title>')
        svg_parts.append('</rect>')
        
        # Texte sur la barre (si assez large)
        if bar_width > 40:
            svg_parts.append(f'<text x="{bar_x + bar_width/2}" y="{y + bar_height/2 + 3}" text-anchor="middle" class="bar-text">{row["duree_jours"]}j</text>')
    
    # Légende
    legend_y = margin_top + chart_height + 40
    legend_x = margin_left
    for i, (cat, color) in enumerate(color_map.items()):
        x_offset = legend_x + (i % 4) * 280
        y_offset = legend_y + (i // 4) * 20
        if y_offset < total_height - 10:
            svg_parts.append(f'<rect x="{x_offset}" y="{y_offset}" width="12" height="12" fill="{color}" rx="2"/>')
            cat_short = cat[:35] + "..." if len(cat) > 35 else cat
            svg_parts.append(f'<text x="{x_offset + 18}" y="{y_offset + 10}" class="legend-text">{cat_short}</text>')
    
    svg_parts.append('</svg>')
    
    return '\n'.join(svg_parts)


def generate_html_report(data, svg_by_category):
    """Génère un rapport HTML complet avec navigation."""
    categories = list(svg_by_category.keys())
    
    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport Gantt - {datetime.now().strftime('%Y-%m-%d')}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; }}
        .nav {{ position: fixed; top: 0; left: 0; right: 0; background: #1f4e79; padding: 1rem; z-index: 100; }}
        .nav-content {{ max-width: 1400px; margin: 0 auto; display: flex; justify-content: space-between; align-items: center; }}
        .nav h1 {{ color: white; font-size: 1.5rem; }}
        .nav-links {{ display: flex; gap: 1rem; flex-wrap: wrap; }}
        .nav-links a {{ color: white; text-decoration: none; padding: 0.5rem 1rem; background: rgba(255,255,255,0.1); border-radius: 4px; font-size: 0.85rem; }}
        .nav-links a:hover {{ background: rgba(255,255,255,0.2); }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 5rem 1rem 2rem; }}
        .slide {{ background: white; border-radius: 8px; padding: 2rem; margin-bottom: 2rem; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        .slide h2 {{ color: #1f4e79; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid #e0e0e0; }}
        .stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 1rem; margin: 1rem 0; }}
        .stat-card {{ background: #f0f4f8; padding: 1rem; border-radius: 8px; text-align: center; }}
        .stat-value {{ font-size: 2rem; font-weight: bold; color: #1f4e79; }}
        .stat-label {{ color: #666; font-size: 0.9rem; }}
        .svg-container {{ overflow-x: auto; margin: 1rem 0; }}
        table {{ width: 100%; border-collapse: collapse; margin: 1rem 0; }}
        th, td {{ padding: 0.75rem; text-align: left; border-bottom: 1px solid #e0e0e0; }}
        th {{ background: #f0f4f8; font-weight: 600; color: #1f4e79; }}
        tr:hover {{ background: #f9f9f9; }}
        @media print {{ .nav {{ display: none; }} .container {{ padding-top: 1rem; }} .slide {{ break-inside: avoid; }} }}
    </style>
</head>
<body>
    <nav class="nav">
        <div class="nav-content">
            <h1>📊 Rapport Gantt</h1>
            <div class="nav-links">
                <a href="#resume">Résumé</a>
"""
    
    for i, cat in enumerate(categories):
        html += f'                <a href="#cat-{i}">{cat[:20]}...</a>\n'
    
    html += """            </div>
        </div>
    </nav>
    <div class="container">
"""
    
    # Slide résumé
    total_tasks = len(data)
    total_categories = len(categories)
    avg_duration = data['duree_jours'].mean()
    min_date = data['debut'].min().strftime('%d/%m/%Y')
    max_date = data['fin'].max().strftime('%d/%m/%Y')
    
    html += f"""
        <div class="slide" id="resume">
            <h2>📈 Résumé Global</h2>
            <div class="stats">
                <div class="stat-card">
                    <div class="stat-value">{total_tasks}</div>
                    <div class="stat-label">Tâches</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">{total_categories}</div>
                    <div class="stat-label">Catégories</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">{avg_duration:.0f}j</div>
                    <div class="stat-label">Durée moyenne</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">{min_date}</div>
                    <div class="stat-label">Date début</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">{max_date}</div>
                    <div class="stat-label">Date fin</div>
                </div>
            </div>
            <h3>Répartition par catégorie</h3>
            <table>
                <thead>
                    <tr><th>Catégorie</th><th>Nb tâches</th><th>Durée moy.</th><th>Période</th></tr>
                </thead>
                <tbody>
"""
    
    for cat in categories:
        cat_data = data[data['categorie'] == cat]
        html += f"""                    <tr>
                        <td>{cat}</td>
                        <td>{len(cat_data)}</td>
                        <td>{cat_data['duree_jours'].mean():.0f} jours</td>
                        <td>{cat_data['debut'].min().strftime('%d/%m/%Y')} → {cat_data['fin'].max().strftime('%d/%m/%Y')}</td>
                    </tr>
"""
    
    html += """                </tbody>
            </table>
        </div>
"""
    
    # Slides par catégorie
    for i, (cat, svg) in enumerate(svg_by_category.items()):
        cat_data = data[data['categorie'] == cat]
        html += f"""
        <div class="slide" id="cat-{i}">
            <h2>{cat}</h2>
            <p>{len(cat_data)} tâches | Durée moyenne: {cat_data['duree_jours'].mean():.0f} jours</p>
            <div class="svg-container">
                {svg}
            </div>
        </div>
"""
    
    html += """
    </div>
</body>
</html>"""
    
    return html


def main():
    """Fonction principale de l'application Streamlit."""
    
    st.markdown('<p class="main-header">📊 Gantt Generic - Générateur de Diagrammes</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Créez des diagrammes de Gantt interactifs à partir de vos fichiers Excel</p>', unsafe_allow_html=True)
    
    # Sidebar - Instructions
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        **Colonnes requises dans votre fichier:**
        1. `catégorie` - Groupe/catégorie de la tâche
        2. `tâche` - Nom de la tâche
        3. `début` - Date de début
        4. `fin` - Date de fin
        
        **Formats de date acceptés:**
        - `YYYY-MM-DD`
        - `DD/MM/YYYY`
        - `DD-MM-YYYY`
        
        **Formats de fichier:**
        - Excel (.xlsx, .xls)
        - CSV (.csv)
        """)
        
        st.divider()
        
        # Téléchargement du template
        st.header("📥 Template")
        template_data = {
            'catégorie': ['Développement', 'Développement', 'Tests', 'Tests', 'Déploiement'],
            'tâche': ['Analyse', 'Codage', 'Tests unitaires', 'Tests intégration', 'Mise en production'],
            'début': ['2025-01-01', '2025-01-15', '2025-02-01', '2025-02-15', '2025-03-01'],
            'fin': ['2025-01-14', '2025-01-31', '2025-02-14', '2025-02-28', '2025-03-15']
        }
        template_df = pd.DataFrame(template_data)
        
        buffer = io.BytesIO()
        template_df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        
        st.download_button(
            label="⬇️ Télécharger le template Excel",
            data=buffer,
            file_name="template_gantt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Zone principale - Upload
    uploaded_file = st.file_uploader(
        "📁 Chargez votre fichier Excel ou CSV",
        type=['xlsx', 'xls', 'csv'],
        help="Le fichier doit contenir les colonnes: catégorie, tâche, début, fin"
    )
    
    if uploaded_file is not None:
        # Chargement des données
        with st.spinner("Analyse du fichier en cours..."):
            data, error = load_and_validate_data(uploaded_file)
        
        if error:
            st.error(f"❌ {error}")
            return
        
        st.success(f"✅ {len(data)} tâches chargées avec succès")
        
        # Métriques
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📋 Tâches", len(data))
        with col2:
            st.metric("📁 Catégories", data['categorie'].nunique())
        with col3:
            st.metric("⏱️ Durée moyenne", f"{data['duree_jours'].mean():.0f}j")
        with col4:
            total_span = (data['fin'].max() - data['debut'].min()).days
            st.metric("📅 Période totale", f"{total_span}j")
        
        # Aperçu des données
        with st.expander("👁️ Aperçu des données", expanded=False):
            st.dataframe(
                data[['categorie', 'tache', 'debut', 'fin', 'duree_jours']].head(20),
                use_container_width=True
            )
        
        # Options de visualisation
        st.divider()
        st.subheader("⚙️ Options de visualisation")
        
        col1, col2 = st.columns(2)
        with col1:
            view_mode = st.radio(
                "Mode d'affichage",
                ["Vue globale", "Par catégorie"],
                horizontal=True
            )
        with col2:
            chart_title = st.text_input("Titre du diagramme", value="Diagramme de Gantt")
        
        # Génération des diagrammes
        st.divider()
        
        if view_mode == "Vue globale":
            st.subheader("📊 Diagramme de Gantt - Vue Globale")
            svg = generate_gantt_svg(data, title=chart_title)
            st.markdown(svg, unsafe_allow_html=True)
        else:
            st.subheader("📊 Diagrammes par Catégorie")
            categories = data['categorie'].unique()
            
            tabs = st.tabs(list(categories))
            for tab, cat in zip(tabs, categories):
                with tab:
                    cat_data = data[data['categorie'] == cat].reset_index(drop=True)
                    svg = generate_gantt_svg(cat_data, title=f"{chart_title} - {cat}")
                    st.markdown(svg, unsafe_allow_html=True)
                    st.caption(f"{len(cat_data)} tâches | Durée moyenne: {cat_data['duree_jours'].mean():.0f} jours")
        
        # Exports
        st.divider()
        st.subheader("📥 Exports")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Export CSV
            csv_buffer = io.StringIO()
            data.to_csv(csv_buffer, index=False)
            st.download_button(
                label="⬇️ Télécharger CSV",
                data=csv_buffer.getvalue(),
                file_name=f"gantt_export_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Export SVG
            svg_global = generate_gantt_svg(data, title=chart_title)
            st.download_button(
                label="⬇️ Télécharger SVG",
                data=svg_global,
                file_name=f"gantt_{datetime.now().strftime('%Y%m%d')}.svg",
                mime="image/svg+xml"
            )
        
        with col3:
            # Export HTML complet
            svg_by_category = {}
            for cat in data['categorie'].unique():
                cat_data = data[data['categorie'] == cat].reset_index(drop=True)
                svg_by_category[cat] = generate_gantt_svg(cat_data, title=cat)
            
            html_report = generate_html_report(data, svg_by_category)
            st.download_button(
                label="⬇️ Télécharger Rapport HTML",
                data=html_report,
                file_name=f"gantt_report_{datetime.now().strftime('%Y%m%d')}.html",
                mime="text/html"
            )
    
    else:
        # État initial - exemple
        st.info("👆 Chargez un fichier Excel ou CSV pour commencer, ou téléchargez le template dans la barre latérale.")
        
        # Démonstration avec données exemple
        with st.expander("🎯 Voir une démonstration", expanded=True):
            demo_data = pd.DataFrame({
                'categorie': ['Phase 1', 'Phase 1', 'Phase 2', 'Phase 2', 'Phase 3'],
                'tache': ['Planification', 'Analyse des besoins', 'Développement', 'Tests', 'Déploiement'],
                'debut': [datetime(2025, 1, 1), datetime(2025, 1, 15), datetime(2025, 2, 1), datetime(2025, 3, 1), datetime(2025, 4, 1)],
                'fin': [datetime(2025, 1, 14), datetime(2025, 1, 31), datetime(2025, 2, 28), datetime(2025, 3, 31), datetime(2025, 4, 15)]
            })
            demo_data['duree_jours'] = (demo_data['fin'] - demo_data['debut']).dt.days + 1
            
            svg_demo = generate_gantt_svg(demo_data, title="Exemple de Projet")
            st.markdown(svg_demo, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
