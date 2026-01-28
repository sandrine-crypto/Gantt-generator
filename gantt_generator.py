#!/usr/bin/env python3
"""
GANTT CHART GENERATOR - Modèles Murins CRUPPE
Génère des diagrammes de Gantt par product line à partir de la feuille "tableau complet"
Exporte en slides HTML interactives et CSV de synthèse

USAGE:
    python gantt_generator.py models-list-20260105-CRUPPE.xlsx
    python gantt_generator.py models-list-20260105-CRUPPE.xlsx --html custom_gantt.html --csv export.csv
    
REQUIREMENTS:
    pandas, openpyxl
    
INSTALLATION:
    pip install pandas openpyxl
"""

import pandas as pd
from datetime import datetime
from pathlib import Path
import argparse
import sys


class GanttGenerator:
    """Générateur de diagrammes de Gantt pour modèles murins"""
    
    def __init__(self, excel_file):
        """Initialise avec un fichier Excel"""
        self.excel_file = excel_file
        self.df = None
        self.df_clean = None
        self.gantt_svgs = {}
        self.min_date = None
        self.max_date = None
        
    def load_data(self):
        """Charge et prépare les données"""
        print(f"📊 Chargement de {self.excel_file}...")
        
        try:
            self.df = pd.read_excel(self.excel_file, sheet_name='tableau complet')
        except Exception as e:
            print(f"❌ Erreur: impossible de lire 'tableau complet' - {e}")
            sys.exit(1)
        
        # Nettoyage des données
        self.df['internal code'] = self.df['internal code'].fillna('N/A')
        self.df['target'] = self.df['target'].fillna('N/A')
        self.df['date disponibilité HO'] = pd.to_datetime(
            self.df['date disponibilité HO'], errors='coerce'
        )
        self.df['data de fin de validation (fin du dernier MI critique taggé validation)'] = pd.to_datetime(
            self.df['data de fin de validation (fin du dernier MI critique taggé validation)'], errors='coerce'
        )
        
        # Filtrer les données valides
        self.df_clean = self.df[
            (self.df['internal code'] != 'N/A') & 
            (self.df['date disponibilité HO'].notna()) & 
            (self.df['data de fin de validation (fin du dernier MI critique taggé validation)'].notna())
        ].copy()
        
        self.df_clean['start_date'] = self.df_clean['date disponibilité HO']
        self.df_clean['end_date'] = self.df_clean['data de fin de validation (fin du dernier MI critique taggé validation)']
        self.df_clean['duration_days'] = (self.df_clean['end_date'] - self.df_clean['start_date']).dt.days
        
        self.min_date = self.df_clean['start_date'].min()
        self.max_date = self.df_clean['end_date'].max()
        
        print(f"✓ {len(self.df_clean)} modèles valides chargés")
        print(f"✓ {self.df_clean['product line'].nunique()} product lines")
        return self
    
    def create_gantt_svg(self, product_line_data, product_line_name):
        """Crée un Gantt chart SVG pour une product line"""
        
        data = product_line_data.sort_values('start_date').reset_index(drop=True)
        date_range = (self.max_date - self.min_date).days
        
        # Paramètres SVG
        bar_height = 35
        left_margin = 260
        top_margin = 100
        total_height = len(data) * bar_height + top_margin + 100
        total_width = 1300
        chart_width = total_width - left_margin - 50
        
        colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', 
                  '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b']
        
        # Construction du SVG
        svg = f'<svg width="{total_width}" height="{total_height}" xmlns="http://www.w3.org/2000/svg">'
        svg += '<defs><style>'
        svg += '.gantt-label { font-size: 12px; font-family: Arial; }'
        svg += '.gantt-label-code { font-weight: bold; font-size: 13px; fill: #2c3e50; }'
        svg += '.gantt-label-target { font-size: 10px; fill: #7f8c8d; }'
        svg += '.gantt-title { font-size: 18px; font-weight: bold; fill: #2c3e50; }'
        svg += '.gantt-bar { opacity: 0.85; stroke: white; stroke-width: 1; }'
        svg += '.gantt-bar-text { font-size: 11px; fill: white; font-weight: bold; text-anchor: middle; }'
        svg += '.gantt-grid-line { stroke: #ecf0f1; stroke-width: 1; }'
        svg += '.gantt-date { font-size: 10px; fill: #7f8c8d; text-anchor: middle; }'
        svg += '.gantt-legend { font-size: 9px; fill: #7f8c8d; }'
        svg += '</style></defs>'
        
        # Fond et titres
        svg += f'<rect width="{total_width}" height="{total_height}" fill="white"/>'
        svg += f'<text x="{left_margin}" y="35" class="gantt-title">{product_line_name}</text>'
        svg += f'<text x="{left_margin}" y="55" class="gantt-legend">({len(data)} modeles)</text>'
        
        # Grille de dates
        step_days = 90
        current = self.min_date
        while current <= self.max_date:
            x_pos = left_margin + (current - self.min_date).days / date_range * chart_width
            svg += f'<line x1="{x_pos}" y1="{top_margin}" x2="{x_pos}" y2="{total_height - 50}" class="gantt-grid-line"/>'
            date_str = current.strftime('%b %Y')
            svg += f'<text x="{x_pos}" y="{top_margin - 15}" class="gantt-date">{date_str}</text>'
            current = pd.Timestamp(current) + pd.Timedelta(days=step_days)
        
        # Barres Gantt
        for idx, (i, row) in enumerate(data.iterrows()):
            y_pos = top_margin + idx * bar_height + 15
            code = str(row['internal code'])[:25]
            target = str(row['target'])[:25]
            
            # Labels de gauche
            svg += f'<text x="10" y="{y_pos + 5}" class="gantt-label gantt-label-code">{code}</text>'
            svg += f'<text x="10" y="{y_pos + 18}" class="gantt-label gantt-label-target">{target}</text>'
            
            # Calcul de la barre
            start_x = left_margin + (row['start_date'] - self.min_date).days / date_range * chart_width
            end_x = left_margin + (row['end_date'] - self.min_date).days / date_range * chart_width
            bar_width = max(end_x - start_x, 3)
            color = colors[idx % len(colors)]
            
            # Infos pour tooltip
            ho_date = row['start_date'].strftime('%d/%m/%Y')
            val_date = row['end_date'].strftime('%d/%m/%Y')
            duration = int(row['duration_days'])
            
            # Rectangle
            svg += f'<rect x="{start_x}" y="{y_pos - 8}" width="{bar_width}" height="25" fill="{color}" class="gantt-bar" title="{code} | {target} | HO: {ho_date} -&gt; Val: {val_date} | {duration}j"/>'
            
            # Texte dans la barre
            if bar_width > 40:
                text_x = start_x + bar_width / 2
                svg += f'<text x="{text_x}" y="{y_pos + 2}" class="gantt-bar-text">{duration}j</text>'
        
        # Axe horizontal et légende
        svg += f'<line x1="{left_margin}" y1="{top_margin + len(data) * bar_height}" x2="{total_width - 50}" y2="{top_margin + len(data) * bar_height}" stroke="#34495e" stroke-width="2"/>'
        svg += f'<text x="{left_margin}" y="{total_height - 30}" class="gantt-legend">Disponibilite HO → Fin Validation</text>'
        svg += '</svg>'
        
        return svg
    
    def generate_gantt_charts(self):
        """Génère tous les Gantt charts"""
        print("\n📈 Génération des Gantt charts...")
        
        product_lines = sorted(self.df_clean['product line'].unique())
        
        for pl in product_lines:
            pl_data = self.df_clean[self.df_clean['product line'] == pl]
            svg = self.create_gantt_svg(pl_data, pl)
            self.gantt_svgs[pl] = svg
            print(f"  ✓ {pl[:50]}... ({len(pl_data)} modèles)")
        
        return self
    
    def generate_html_slides(self, output_file='gantt_slides.html'):
        """Génère les slides HTML avec les Gantt charts"""
        print(f"\n📝 Génération des slides HTML ({output_file})...")
        
        html = """<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gantt Charts - Modeles Murins CRUPPE</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .slide-container {
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            margin: 20px auto;
            padding: 30px;
            max-width: 1400px;
            page-break-after: always;
        }
        .slide-header {
            border-bottom: 3px solid #667eea;
            padding-bottom: 15px;
            margin-bottom: 20px;
        }
        .slide-title {
            font-size: 28px;
            color: #2c3e50;
            font-weight: bold;
            margin-bottom: 5px;
        }
        .slide-subtitle {
            font-size: 14px;
            color: #7f8c8d;
        }
        .gantt-wrapper {
            overflow-x: auto;
            margin-top: 20px;
        }
        .gantt-wrapper svg {
            display: block;
            margin: 0 auto;
        }
        @media print {
            .slide-container { page-break-after: always; margin: 0; }
        }
    </style>
</head>
<body>
"""
        
        # Slide de couverture
        html += f"""
<div class="slide-container">
    <div class="slide-header">
        <div class="slide-title">Diagrammes de Gantt</div>
        <div class="slide-subtitle">Modeles Murins Genetiquement Modifies - CRUPPE</div>
    </div>
    <div style="text-align: center; margin-top: 60px;">
        <p style="font-size: 18px; color: #7f8c8d; margin: 20px 0;">
            <strong>{len(self.df_clean)} modeles</strong> repartis dans <strong>{len(self.gantt_svgs)} product lines</strong>
        </p>
        <p style="font-size: 14px; color: #95a5a6; margin: 20px 0;">
            Periode: {self.min_date.strftime('%d/%m/%Y')} → {self.max_date.strftime('%d/%m/%Y')}
        </p>
        <p style="font-size: 12px; color: #bdc3c7; margin: 40px 0;">
            Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}
        </p>
    </div>
</div>
"""
        
        # Slides Gantt
        for idx, (product_line, svg) in enumerate(self.gantt_svgs.items(), 1):
            html += f"""
<div class="slide-container">
    <div class="slide-header">
        <div class="slide-title">Produit {idx}/{len(self.gantt_svgs)}</div>
        <div class="slide-subtitle">{product_line}</div>
    </div>
    <div class="gantt-wrapper">
        {svg}
    </div>
</div>
"""
        
        # Slide de résumé
        html += """
<div class="slide-container">
    <div class="slide-header">
        <div class="slide-title">Resume</div>
        <div class="slide-subtitle">Vue d'ensemble des donnees</div>
    </div>
    <div style="margin-top: 40px;">
        <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr style="background: #ecf0f1;">
                <th style="border: 1px solid #bdc3c7; padding: 12px; text-align: left;">Product Line</th>
                <th style="border: 1px solid #bdc3c7; padding: 12px; text-align: center;">Modeles</th>
            </tr>
"""
        
        for product_line, count in sorted(
            self.df_clean['product line'].value_counts().items(), 
            key=lambda x: x[1], 
            reverse=True
        ):
            html += f"""
            <tr>
                <td style="border: 1px solid #bdc3c7; padding: 12px;">{product_line}</td>
                <td style="border: 1px solid #bdc3c7; padding: 12px; text-align: center;">{count}</td>
            </tr>
"""
        
        html += """
        </table>
    </div>
</div>

</body>
</html>
"""
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
        
        print(f"✓ Slides générées: {output_file}")
        print(f"  - 1 slide couverture + {len(self.gantt_svgs)} slides Gantt + 1 résumé = {len(self.gantt_svgs) + 2} slides totales")
        return self
    
    def export_csv(self, output_file='gantt_models_export.csv'):
        """Exporte les données en CSV"""
        print(f"\n📊 Export CSV ({output_file})...")
        
        export_df = self.df_clean[[
            'product line', 'internal code', 'target', 'status',
            'start_date', 'end_date', 'duration_days'
        ]].copy()
        
        export_df.columns = [
            'Product Line', 'Internal Code', 'Target', 'Status',
            'Date Disponibilite HO', 'Date Fin Validation', 'Duree (jours)'
        ]
        
        export_df = export_df.sort_values('Product Line')
        export_df.to_csv(output_file, index=False, encoding='utf-8')
        
        print(f"✓ CSV exporté: {output_file} ({len(export_df)} lignes)")
        return self


def main():
    parser = argparse.ArgumentParser(
        description='Générateur de diagrammes de Gantt pour modèles murins',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
EXEMPLES:
  python gantt_generator.py models-list.xlsx
  python gantt_generator.py models-list.xlsx --html custom.html --csv export.csv
        """
    )
    parser.add_argument('excel_file', help='Chemin du fichier Excel')
    parser.add_argument('--html', default='gantt_slides.html', help='Fichier HTML de sortie')
    parser.add_argument('--csv', default='gantt_models_export.csv', help='Fichier CSV de sortie')
    
    args = parser.parse_args()
    
    if not Path(args.excel_file).exists():
        print(f"❌ Erreur: {args.excel_file} non trouvé")
        sys.exit(1)
    
    try:
        print("\n" + "="*60)
        print("GANTT CHART GENERATOR - Modèles Murins CRUPPE")
        print("="*60 + "\n")
        
        generator = GanttGenerator(args.excel_file)
        generator.load_data().generate_gantt_charts().generate_html_slides(args.html).export_csv(args.csv)
        
        print(f"\n{'='*60}")
        print("✅ SUCCES! Fichiers générés:")
        print(f"   • {args.html} (slides HTML interactives)")
        print(f"   • {args.csv} (données d'export)")
        print(f"{'='*60}\n")
        
    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
