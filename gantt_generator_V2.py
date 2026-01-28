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

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

import pandas as pd

if TYPE_CHECKING:
    from pandas import DataFrame, Timestamp


# =============================================================================
# CONSTANTES DE CONFIGURATION
# =============================================================================

@dataclass(frozen=True)
class GanttConfig:
    """Configuration centralisée pour les Gantt charts."""

    # Dimensions SVG
    BAR_HEIGHT: int = 35
    LEFT_MARGIN: int = 260
    TOP_MARGIN: int = 100
    TOTAL_WIDTH: int = 1300
    RIGHT_MARGIN: int = 50
    BOTTOM_MARGIN: int = 100
    MIN_BAR_WIDTH: int = 3
    MIN_BAR_WIDTH_FOR_TEXT: int = 40

    # Grille temporelle
    GRID_STEP_DAYS: int = 90

    # Palette de couleurs
    COLORS: tuple[str, ...] = (
        '#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6',
        '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b'
    )

    # Limites de texte
    MAX_CODE_LENGTH: int = 25
    MAX_TARGET_LENGTH: int = 25


CONFIG = GanttConfig()

# Styles CSS réutilisables (définis une seule fois)
SVG_STYLES = """
.gantt-label { font-size: 12px; font-family: Arial; }
.gantt-label-code { font-weight: bold; font-size: 13px; fill: #2c3e50; }
.gantt-label-target { font-size: 10px; fill: #7f8c8d; }
.gantt-title { font-size: 18px; font-weight: bold; fill: #2c3e50; }
.gantt-bar { opacity: 0.85; stroke: white; stroke-width: 1; }
.gantt-bar-text { font-size: 11px; fill: white; font-weight: bold; text-anchor: middle; }
.gantt-grid-line { stroke: #ecf0f1; stroke-width: 1; }
.gantt-date { font-size: 10px; fill: #7f8c8d; text-anchor: middle; }
.gantt-legend { font-size: 9px; fill: #7f8c8d; }
""".strip()

HTML_STYLES = """
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
.slide-title { font-size: 28px; color: #2c3e50; font-weight: bold; margin-bottom: 5px; }
.slide-subtitle { font-size: 14px; color: #7f8c8d; }
.gantt-wrapper { overflow-x: auto; margin-top: 20px; }
.gantt-wrapper svg { display: block; margin: 0 auto; }
.summary-table { width: 100%; border-collapse: collapse; font-size: 14px; }
.summary-table th, .summary-table td { border: 1px solid #bdc3c7; padding: 12px; }
.summary-table th { background: #ecf0f1; text-align: left; }
.summary-table td:last-child { text-align: center; }
@media print { .slide-container { page-break-after: always; margin: 0; } }
""".strip()


# =============================================================================
# CLASSE PRINCIPALE
# =============================================================================

class GanttGenerator:
    """Générateur de diagrammes de Gantt pour modèles murins."""

    __slots__ = ('excel_file', 'df', 'df_clean', 'gantt_svgs', 'min_date', 'max_date', '_date_range_days')

    def __init__(self, excel_file: str | Path) -> None:
        """Initialise avec un fichier Excel."""
        self.excel_file = Path(excel_file)
        self.df: DataFrame | None = None
        self.df_clean: DataFrame | None = None
        self.gantt_svgs: dict[str, str] = {}
        self.min_date: Timestamp | None = None
        self.max_date: Timestamp | None = None
        self._date_range_days: int = 0

    def load_data(self) -> GanttGenerator:
        """Charge et prépare les données."""
        print(f"📊 Chargement de {self.excel_file}...")

        try:
            self.df = pd.read_excel(self.excel_file, sheet_name='tableau complet')
        except Exception as e:
            print(f"❌ Erreur: impossible de lire 'tableau complet' - {e}")
            sys.exit(1)

        # Colonnes source
        col_start = 'date disponibilité HO'
        col_end = 'data de fin de validation (fin du dernier MI critique taggé validation)'

        # Nettoyage vectorisé des données
        df = self.df.copy()
        df['internal code'] = df['internal code'].fillna('N/A')
        df['target'] = df['target'].fillna('N/A')
        df[col_start] = pd.to_datetime(df[col_start], errors='coerce')
        df[col_end] = pd.to_datetime(df[col_end], errors='coerce')

        # Filtrage avec masque booléen (plus efficace)
        mask = (
            (df['internal code'] != 'N/A') &
            df[col_start].notna() &
            df[col_end].notna()
        )
        self.df_clean = df.loc[mask].copy()

        # Colonnes de travail
        self.df_clean['start_date'] = self.df_clean[col_start]
        self.df_clean['end_date'] = self.df_clean[col_end]
        self.df_clean['duration_days'] = (
            self.df_clean['end_date'] - self.df_clean['start_date']
        ).dt.days

        # Calcul unique des bornes temporelles
        self.min_date = self.df_clean['start_date'].min()
        self.max_date = self.df_clean['end_date'].max()
        self._date_range_days = (self.max_date - self.min_date).days

        print(f"✓ {len(self.df_clean)} modèles valides chargés")
        print(f"✓ {self.df_clean['product line'].nunique()} product lines")
        return self

    def _compute_x_position(self, date: Timestamp, chart_width: float) -> float:
        """Calcule la position X pour une date donnée."""
        days_from_start = (date - self.min_date).days
        return CONFIG.LEFT_MARGIN + (days_from_start / self._date_range_days) * chart_width

    def _build_svg_header(self, total_width: int, total_height: int,
                          product_line_name: str, model_count: int) -> list[str]:
        """Construit l'en-tête SVG avec styles et titre."""
        return [
            f'<svg width="{total_width}" height="{total_height}" xmlns="http://www.w3.org/2000/svg">',
            f'<defs><style>{SVG_STYLES}</style></defs>',
            f'<rect width="{total_width}" height="{total_height}" fill="white"/>',
            f'<text x="{CONFIG.LEFT_MARGIN}" y="35" class="gantt-title">{product_line_name}</text>',
            f'<text x="{CONFIG.LEFT_MARGIN}" y="55" class="gantt-legend">({model_count} modeles)</text>',
        ]

    def _build_date_grid(self, chart_width: float, total_height: int) -> list[str]:
        """Construit la grille de dates avec pd.date_range (optimisé)."""
        parts: list[str] = []

        # Utilisation de pd.date_range au lieu d'une boucle while
        date_range = pd.date_range(
            start=self.min_date,
            end=self.max_date,
            freq=f'{CONFIG.GRID_STEP_DAYS}D'
        )

        for current in date_range:
            x_pos = self._compute_x_position(current, chart_width)
            date_str = current.strftime('%b %Y')
            parts.append(
                f'<line x1="{x_pos}" y1="{CONFIG.TOP_MARGIN}" '
                f'x2="{x_pos}" y2="{total_height - 50}" class="gantt-grid-line"/>'
            )
            parts.append(
                f'<text x="{x_pos}" y="{CONFIG.TOP_MARGIN - 15}" class="gantt-date">{date_str}</text>'
            )

        return parts

    def _build_gantt_bar(self, row: pd.Series, idx: int, chart_width: float) -> list[str]:
        """Construit une barre Gantt avec son label."""
        parts: list[str] = []
        y_pos = CONFIG.TOP_MARGIN + idx * CONFIG.BAR_HEIGHT + 15

        # Troncature des textes
        code = str(row['internal code'])[:CONFIG.MAX_CODE_LENGTH]
        target = str(row['target'])[:CONFIG.MAX_TARGET_LENGTH]

        # Labels de gauche
        parts.append(f'<text x="10" y="{y_pos + 5}" class="gantt-label gantt-label-code">{code}</text>')
        parts.append(f'<text x="10" y="{y_pos + 18}" class="gantt-label gantt-label-target">{target}</text>')

        # Calcul des positions (réutilise la méthode centralisée)
        start_x = self._compute_x_position(row['start_date'], chart_width)
        end_x = self._compute_x_position(row['end_date'], chart_width)
        bar_width = max(end_x - start_x, CONFIG.MIN_BAR_WIDTH)
        color = CONFIG.COLORS[idx % len(CONFIG.COLORS)]

        # Données pour tooltip
        ho_date = row['start_date'].strftime('%d/%m/%Y')
        val_date = row['end_date'].strftime('%d/%m/%Y')
        duration = int(row['duration_days'])

        # Rectangle de la barre
        tooltip = f"{code} | {target} | HO: {ho_date} -&gt; Val: {val_date} | {duration}j"
        parts.append(
            f'<rect x="{start_x}" y="{y_pos - 8}" width="{bar_width}" height="25" '
            f'fill="{color}" class="gantt-bar" title="{tooltip}"/>'
        )

        # Texte dans la barre (si assez large)
        if bar_width > CONFIG.MIN_BAR_WIDTH_FOR_TEXT:
            text_x = start_x + bar_width / 2
            parts.append(f'<text x="{text_x}" y="{y_pos + 2}" class="gantt-bar-text">{duration}j</text>')

        return parts

    def create_gantt_svg(self, product_line_data: DataFrame, product_line_name: str) -> str:
        """Crée un Gantt chart SVG pour une product line."""
        data = product_line_data.sort_values('start_date').reset_index(drop=True)
        model_count = len(data)

        # Dimensions calculées
        chart_width = CONFIG.TOTAL_WIDTH - CONFIG.LEFT_MARGIN - CONFIG.RIGHT_MARGIN
        total_height = model_count * CONFIG.BAR_HEIGHT + CONFIG.TOP_MARGIN + CONFIG.BOTTOM_MARGIN

        # Construction avec liste (O(n) vs O(n²) pour concaténation)
        svg_parts: list[str] = []

        # En-tête
        svg_parts.extend(self._build_svg_header(CONFIG.TOTAL_WIDTH, total_height, product_line_name, model_count))

        # Grille de dates
        svg_parts.extend(self._build_date_grid(chart_width, total_height))

        # Barres Gantt (itération sur les lignes)
        for idx, row in data.iterrows():
            svg_parts.extend(self._build_gantt_bar(row, idx, chart_width))

        # Axe horizontal et légende
        axis_y = CONFIG.TOP_MARGIN + model_count * CONFIG.BAR_HEIGHT
        svg_parts.append(
            f'<line x1="{CONFIG.LEFT_MARGIN}" y1="{axis_y}" '
            f'x2="{CONFIG.TOTAL_WIDTH - CONFIG.RIGHT_MARGIN}" y2="{axis_y}" '
            f'stroke="#34495e" stroke-width="2"/>'
        )
        svg_parts.append(
            f'<text x="{CONFIG.LEFT_MARGIN}" y="{total_height - 30}" class="gantt-legend">'
            f'Disponibilite HO → Fin Validation</text>'
        )
        svg_parts.append('</svg>')

        return ''.join(svg_parts)

    def generate_gantt_charts(self) -> GanttGenerator:
        """Génère tous les Gantt charts."""
        print("\n📈 Génération des Gantt charts...")

        # Groupby est plus efficace que filtrer en boucle
        grouped = self.df_clean.groupby('product line', sort=True)

        for pl_name, pl_data in grouped:
            self.gantt_svgs[pl_name] = self.create_gantt_svg(pl_data, pl_name)
            print(f"  ✓ {pl_name[:50]}... ({len(pl_data)} modèles)")

        return self

    def _build_cover_slide(self) -> str:
        """Construit la slide de couverture."""
        return f"""
<div class="slide-container">
    <div class="slide-header">
        <div class="slide-title">Diagrammes de Gantt</div>
        <div class="slide-subtitle">Modeles Murins Genetiquement Modifies - CRUPPE</div>
    </div>
    <div style="text-align: center; margin-top: 60px;">
        <p style="font-size: 18px; color: #7f8c8d; margin: 20px 0;">
            <strong>{len(self.df_clean)} modeles</strong> repartis dans
            <strong>{len(self.gantt_svgs)} product lines</strong>
        </p>
        <p style="font-size: 14px; color: #95a5a6; margin: 20px 0;">
            Periode: {self.min_date.strftime('%d/%m/%Y')} → {self.max_date.strftime('%d/%m/%Y')}
        </p>
        <p style="font-size: 12px; color: #bdc3c7; margin: 40px 0;">
            Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}
        </p>
    </div>
</div>"""

    def _build_gantt_slide(self, idx: int, product_line: str, svg: str) -> str:
        """Construit une slide Gantt."""
        return f"""
<div class="slide-container">
    <div class="slide-header">
        <div class="slide-title">Produit {idx}/{len(self.gantt_svgs)}</div>
        <div class="slide-subtitle">{product_line}</div>
    </div>
    <div class="gantt-wrapper">{svg}</div>
</div>"""

    def _build_summary_slide(self) -> str:
        """Construit la slide de résumé."""
        # Pré-calcul des comptages (une seule passe)
        counts = self.df_clean['product line'].value_counts().sort_values(ascending=False)

        rows = ''.join(
            f'<tr><td>{pl}</td><td>{count}</td></tr>'
            for pl, count in counts.items()
        )

        return f"""
<div class="slide-container">
    <div class="slide-header">
        <div class="slide-title">Resume</div>
        <div class="slide-subtitle">Vue d'ensemble des donnees</div>
    </div>
    <div style="margin-top: 40px;">
        <table class="summary-table">
            <tr><th>Product Line</th><th>Modeles</th></tr>
            {rows}
        </table>
    </div>
</div>"""

    def generate_html_slides(self, output_file: str = 'gantt_slides.html') -> GanttGenerator:
        """Génère les slides HTML avec les Gantt charts."""
        print(f"\n📝 Génération des slides HTML ({output_file})...")

        # Construction avec liste (optimisé)
        html_parts: list[str] = [
            '<!DOCTYPE html>',
            '<html lang="fr">',
            '<head>',
            '<meta charset="UTF-8">',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
            '<title>Gantt Charts - Modeles Murins CRUPPE</title>',
            f'<style>{HTML_STYLES}</style>',
            '</head>',
            '<body>',
        ]

        # Slides
        html_parts.append(self._build_cover_slide())

        for idx, (product_line, svg) in enumerate(self.gantt_svgs.items(), 1):
            html_parts.append(self._build_gantt_slide(idx, product_line, svg))

        html_parts.append(self._build_summary_slide())
        html_parts.extend(['</body>', '</html>'])

        # Écriture unique
        Path(output_file).write_text(''.join(html_parts), encoding='utf-8')

        total_slides = len(self.gantt_svgs) + 2
        print(f"✓ Slides générées: {output_file}")
        print(f"  - 1 slide couverture + {len(self.gantt_svgs)} slides Gantt + 1 résumé = {total_slides} slides totales")
        return self

    def export_csv(self, output_file: str = 'gantt_models_export.csv') -> GanttGenerator:
        """Exporte les données en CSV."""
        print(f"\n📊 Export CSV ({output_file})...")

        # Sélection et renommage en une seule opération
        column_mapping = {
            'product line': 'Product Line',
            'internal code': 'Internal Code',
            'target': 'Target',
            'status': 'Status',
            'start_date': 'Date Disponibilite HO',
            'end_date': 'Date Fin Validation',
            'duration_days': 'Duree (jours)',
        }

        export_df = (
            self.df_clean[list(column_mapping.keys())]
            .rename(columns=column_mapping)
            .sort_values('Product Line')
        )

        export_df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"✓ CSV exporté: {output_file} ({len(export_df)} lignes)")
        return self


# =============================================================================
# POINT D'ENTRÉE
# =============================================================================

def main() -> None:
    """Point d'entrée principal."""
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

    excel_path = Path(args.excel_file)
    if not excel_path.exists():
        print(f"❌ Erreur: {args.excel_file} non trouvé")
        sys.exit(1)

    try:
        print("\n" + "=" * 60)
        print("GANTT CHART GENERATOR - Modèles Murins CRUPPE")
        print("=" * 60 + "\n")

        (
            GanttGenerator(excel_path)
            .load_data()
            .generate_gantt_charts()
            .generate_html_slides(args.html)
            .export_csv(args.csv)
        )

        print(f"\n{'=' * 60}")
        print("✅ SUCCES! Fichiers générés:")
        print(f"   • {args.html} (slides HTML interactives)")
        print(f"   • {args.csv} (données d'export)")
        print(f"{'=' * 60}\n")

    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
