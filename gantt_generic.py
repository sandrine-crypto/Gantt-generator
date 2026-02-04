"""
Gantt Generic - Application Streamlit
Génère des diagrammes de Gantt interactifs à partir de fichiers Excel
Colonnes requises: catégorie, tâche, début, fin
Exports: HTML, PowerPoint (PPTX), Word (DOCX), CSV, SVG
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import html
import subprocess
import tempfile
import os
import json

st.set_page_config(page_title="Gantt Generic", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f4e79; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #666; margin-bottom: 2rem; }
    .stDownloadButton > button { width: 100%; }
    .gantt-container { overflow-x: auto; overflow-y: auto; max-height: 80vh; }
</style>
""", unsafe_allow_html=True)


def parse_date(date_val):
    if pd.isna(date_val): return None
    if isinstance(date_val, datetime): return date_val
    if isinstance(date_val, pd.Timestamp): return date_val.to_pydatetime()
    if isinstance(date_val, str):
        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y"]:
            try: return datetime.strptime(date_val.strip(), fmt)
            except: pass
    return None


def find_column(df, names):
    cols = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower() in cols: return cols[n.lower()]
    return None


def load_data(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, engine='openpyxl')
        mapping = {
            'categorie': find_column(df, ['catégorie', 'categorie', 'category', 'groupe', 'group']),
            'tache': find_column(df, ['tâche', 'tache', 'task', 'nom', 'name', 'activité']),
            'debut': find_column(df, ['début', 'debut', 'start', 'date_debut', 'start_date']),
            'fin': find_column(df, ['fin', 'end', 'date_fin', 'end_date', 'échéance'])
        }
        missing = [k for k, v in mapping.items() if not v]
        if missing: return None, f"Colonnes manquantes: {', '.join(missing)}"
        
        data = pd.DataFrame({
            'categorie': df[mapping['categorie']].astype(str).str.strip(),
            'tache': df[mapping['tache']].astype(str).str.strip(),
            'debut': df[mapping['debut']].apply(parse_date),
            'fin': df[mapping['fin']].apply(parse_date)
        })
        data = data.dropna(subset=['debut', 'fin'])
        data = data[data['tache'].notna() & (data['tache'] != '') & (data['tache'] != 'nan')]
        data['duree_jours'] = (data['fin'] - data['debut']).dt.days + 1
        data = data.sort_values(['categorie', 'debut']).reset_index(drop=True)
        return (data, None) if len(data) > 0 else (None, "Aucune donnée valide")
    except Exception as e:
        return None, str(e)


def colors(n):
    palette = ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac"]
    return [palette[i % len(palette)] for i in range(n)]


def esc(t): return html.escape(str(t))


def generate_svg(data, title="Diagramme de Gantt"):
    if len(data) == 0: return "<svg><text>Aucune donnée</text></svg>"
    
    ml = min(500, max(250, data['tache'].str.len().max() * 8 + 50))
    mr, mt, mb = 50, 100, 80
    lh, bh, rs = 18, 24, 12
    rh = lh + bh + rs
    
    n = len(data)
    ch = n * rh
    cw = 800
    tw, th = ml + cw + mr, mt + ch + mb
    
    mind, maxd = data['debut'].min(), data['fin'].max()
    dr = max(1, (maxd - mind).days)
    
    cats = data['categorie'].unique()
    cmap = dict(zip(cats, colors(len(cats))))
    
    s = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {tw} {th}" width="{tw}" height="{th}" style="font-family:Arial,sans-serif;">']
    s.append('<defs><style>.t{font-size:22px;font-weight:bold;fill:#1f4e79}.st{font-size:13px;fill:#666}.ax{font-size:11px;fill:#555}.tl{font-size:12px;font-weight:500;fill:#333}.ta{font-size:11px;fill:#444}.dt{font-size:10px;font-weight:bold;fill:#fff}.gl{stroke:#e8e8e8;stroke-width:1}.bar{rx:4;ry:4}.lt{font-size:11px;fill:#555}</style></defs>')
    s.append(f'<rect width="{tw}" height="{th}" fill="#fafafa"/><rect x="{ml}" y="{mt}" width="{cw}" height="{ch}" fill="#fff" stroke="#ddd"/>')
    s.append(f'<text x="{tw/2}" y="35" text-anchor="middle" class="t">{esc(title)}</text>')
    s.append(f'<text x="{tw/2}" y="58" text-anchor="middle" class="st">Période: {mind.strftime("%d/%m/%Y")} → {maxd.strftime("%d/%m/%Y")} ({dr}j)</text>')
    s.append(f'<text x="{tw/2}" y="78" text-anchor="middle" class="st">{n} tâches | {len(cats)} catégories</text>')
    
    ng = min(10, max(4, dr // 30))
    for i in range(ng + 1):
        x = ml + (i / ng) * cw
        s.append(f'<line x1="{x}" y1="{mt}" x2="{x}" y2="{mt + ch}" class="gl"/>')
        gd = mind + timedelta(days=int(i * dr / ng))
        s.append(f'<text x="{x}" y="{mt + ch + 18}" text-anchor="middle" class="ax">{gd.strftime("%d/%m/%y")}</text>')
    
    cc = None
    for idx, (_, r) in enumerate(data.iterrows()):
        ry = mt + idx * rh
        ly, by = ry + lh - 4, ry + lh + 2
        s.append(f'<line x1="{ml}" y1="{ry + rh}" x2="{ml + cw}" y2="{ry + rh}" class="gl"/>')
        
        so = (r['debut'] - mind).days
        du = (r['fin'] - r['debut']).days + 1
        bx = ml + (so / dr) * cw
        bw = max(8, (du / dr) * cw)
        col = cmap[r['categorie']]
        
        if r['categorie'] != cc:
            cc = r['categorie']
            s.append(f'<rect x="5" y="{ly - 10}" width="8" height="8" fill="{col}" rx="2"/>')
        
        s.append(f'<text x="{ml - 12}" y="{by + bh/2 + 4}" text-anchor="end" class="tl">{esc(r["tache"])}</text>')
        s.append(f'<text x="{bx + 4}" y="{ly}" class="ta">{esc(r["categorie"])} | {r["debut"].strftime("%d/%m")} → {r["fin"].strftime("%d/%m/%y")}</text>')
        
        tip = f'{r["tache"]}\n{r["categorie"]}\n{r["debut"].strftime("%d/%m/%Y")} → {r["fin"].strftime("%d/%m/%Y")}\n{r["duree_jours"]}j'
        s.append(f'<rect x="{bx}" y="{by}" width="{bw}" height="{bh}" fill="{col}" class="bar"><title>{esc(tip)}</title></rect>')
        if bw > 35:
            s.append(f'<text x="{bx + bw/2}" y="{by + bh/2 + 4}" text-anchor="middle" class="dt">{r["duree_jours"]}j</text>')
    
    ly = mt + ch + 40
    s.append(f'<text x="{ml}" y="{ly}" style="font-size:12px;font-weight:bold;fill:#333">Légende:</text>')
    for i, (c, col) in enumerate(cmap.items()):
        xo, yo = ml + (i % 3) * 280, ly + 18 + (i // 3) * 22
        if yo < th - 10:
            s.append(f'<rect x="{xo}" y="{yo - 9}" width="14" height="14" fill="{col}" rx="3"/><text x="{xo + 20}" y="{yo + 2}" class="lt">{esc(c)}</text>')
    
    s.append('</svg>')
    return '\n'.join(s)


def generate_html(data, svg_by_cat, title="Rapport Gantt"):
    cats = list(svg_by_cat.keys())
    svg_all = generate_svg(data, "Vue d'ensemble")
    n, avg = len(data), data['duree_jours'].mean()
    mind, maxd = data['debut'].min().strftime('%d/%m/%Y'), data['fin'].max().strftime('%d/%m/%Y')
    span = (data['fin'].max() - data['debut'].min()).days
    
    h = f'''<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>{esc(title)}</title>
<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:'Segoe UI',sans-serif;background:#f5f7fa}}
.nav{{position:fixed;top:0;left:0;right:0;background:linear-gradient(135deg,#1f4e79,#2d6da3);padding:1rem;z-index:100}}
.nav h1{{color:#fff;font-size:1.3rem}}.container{{max-width:1400px;margin:0 auto;padding:5rem 1rem 2rem}}
.slide{{background:#fff;border-radius:10px;padding:2rem;margin-bottom:2rem;box-shadow:0 2px 10px rgba(0,0,0,.1)}}
.slide h2{{color:#1f4e79;border-bottom:2px solid #e0e0e0;padding-bottom:.5rem;margin-bottom:1rem}}
.stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:1rem;margin:1rem 0}}
.stat{{background:#f0f4f8;padding:1rem;border-radius:8px;text-align:center}}
.stat-value{{font-size:1.8rem;font-weight:bold;color:#1f4e79}}.stat-label{{font-size:.85rem;color:#666}}
.svg-container{{overflow-x:auto;margin:1rem 0}}table{{width:100%;border-collapse:collapse}}
th,td{{padding:.75rem;text-align:left;border-bottom:1px solid #e0e0e0}}th{{background:#f0f4f8;color:#1f4e79}}
@media print{{.nav{{display:none}}.container{{padding-top:1rem}}}}</style></head>
<body><nav class="nav"><h1>📊 {esc(title)}</h1></nav><div class="container">
<div class="slide"><h2>📈 Résumé</h2><div class="stats">
<div class="stat"><div class="stat-value">{n}</div><div class="stat-label">Tâches</div></div>
<div class="stat"><div class="stat-value">{len(cats)}</div><div class="stat-label">Catégories</div></div>
<div class="stat"><div class="stat-value">{avg:.0f}j</div><div class="stat-label">Durée moy.</div></div>
<div class="stat"><div class="stat-value">{span}j</div><div class="stat-label">Période</div></div></div>
<p style="color:#666;margin:1rem 0">Période: {mind} → {maxd}</p>
<table><thead><tr><th>Catégorie</th><th>Tâches</th><th>Durée moy.</th></tr></thead><tbody>'''
    
    for c in cats:
        cd = data[data['categorie'] == c]
        h += f'<tr><td>{esc(c)}</td><td>{len(cd)}</td><td>{cd["duree_jours"].mean():.0f}j</td></tr>'
    
    h += f'</tbody></table></div><div class="slide"><h2>🗓️ Vue d\'ensemble</h2><div class="svg-container">{svg_all}</div></div>'
    
    for c, svg in svg_by_cat.items():
        cd = data[data['categorie'] == c]
        h += f'<div class="slide"><h2>{esc(c)}</h2><p style="color:#666;margin-bottom:1rem">{len(cd)} tâches | Durée moy.: {cd["duree_jours"].mean():.0f}j</p><div class="svg-container">{svg}</div></div>'
    
    return h + '</div></body></html>'


def gen_pptx(data, title):
    cats = data['categorie'].unique()
    cmap = {c: colors(len(cats))[i].replace("#", "") for i, c in enumerate(cats)}
    mind, maxd = data['debut'].min(), data['fin'].max()
    dr = max(1, (maxd - mind).days)
    n, avg = len(data), data['duree_jours'].mean()
    
    js = f'''const pptxgen=require("pptxgenjs");let p=new pptxgen();p.layout="LAYOUT_16x9";
let s1=p.addSlide();s1.background={{color:"1F4E79"}};
s1.addText({json.dumps(title)},{{x:.5,y:2,w:9,h:1,fontSize:40,color:"FFFFFF",bold:true,align:"center"}});
s1.addText("Généré le {datetime.now().strftime('%d/%m/%Y')}",{{x:.5,y:3.2,w:9,h:.5,fontSize:16,color:"CADCFC",align:"center"}});
s1.addText("{n} tâches | {len(cats)} catégories | {avg:.0f}j durée moyenne",{{x:.5,y:4,w:9,h:.4,fontSize:14,color:"CADCFC",align:"center"}});
let s2=p.addSlide();s2.addText("📈 Résumé",{{x:.5,y:.3,w:9,h:.6,fontSize:28,color:"1F4E79",bold:true}});
s2.addText("Période: {mind.strftime('%d/%m/%Y')} → {maxd.strftime('%d/%m/%Y')}",{{x:.5,y:1,w:9,h:.4,fontSize:14,color:"666666"}});
let t=[[{{text:"Catégorie",options:{{bold:true,fill:{{color:"1F4E79"}},color:"FFFFFF"}}}},{{text:"Tâches",options:{{bold:true,fill:{{color:"1F4E79"}},color:"FFFFFF"}}}},{{text:"Durée",options:{{bold:true,fill:{{color:"1F4E79"}},color:"FFFFFF"}}}}]];
'''
    for c in cats:
        cd = data[data['categorie'] == c]
        js += f't.push([{{text:{json.dumps(c)}}},{{text:"{len(cd)}"}},{{text:"{cd["duree_jours"].mean():.0f}j"}}]);'
    
    js += 's2.addTable(t,{x:.5,y:1.5,w:9,h:3,fontSize:11,border:{pt:.5,color:"CCCCCC"}});'
    
    for ci, c in enumerate(cats):
        cd = data[data['categorie'] == c].reset_index(drop=True)
        col = cmap[c]
        js += f'let s{ci+3}=p.addSlide();s{ci+3}.addShape(p.shapes.RECTANGLE,{{x:0,y:0,w:.12,h:5.625,fill:{{color:"{col}"}}}});'
        js += f's{ci+3}.addText({json.dumps(c)},{{x:.3,y:.2,w:9,h:.5,fontSize:22,color:"1F4E79",bold:true}});'
        js += f's{ci+3}.addText("{len(cd)} tâches | Durée moy.: {cd["duree_jours"].mean():.0f}j",{{x:.3,y:.65,w:9,h:.3,fontSize:11,color:"666666"}});'
        
        mx = min(12, len(cd))
        bh = min(.32, 3.8 / mx) if mx > 0 else .32
        for i, (_, r) in enumerate(cd.head(mx).iterrows()):
            ty = 1.1 + i * (bh + .08)
            so = (r['debut'] - mind).days
            du = (r['fin'] - r['debut']).days + 1
            bx = 2.8 + (so / dr) * 6.5
            bw = max(.15, (du / dr) * 6.5)
            tn = r['tache'][:35].replace('"', "'").replace('\\', '')
            js += f's{ci+3}.addText("{tn}",{{x:.3,y:{ty},w:2.4,h:{bh},fontSize:9,color:"333333",valign:"middle"}});'
            js += f's{ci+3}.addShape(p.shapes.RECTANGLE,{{x:{bx},y:{ty},w:{bw},h:{bh},fill:{{color:"{col}"}}}});'
            js += f's{ci+3}.addText("{r["duree_jours"]}j",{{x:{bx},y:{ty},w:{bw},h:{bh},fontSize:8,color:"FFFFFF",bold:true,align:"center",valign:"middle"}});'
    
    js += 'p.writeFile({fileName:"output.pptx"}).then(()=>console.log("OK"));'
    return js


def gen_docx(data, title):
    cats = data['categorie'].unique()
    mind, maxd = data['debut'].min(), data['fin'].max()
    n, avg = len(data), data['duree_jours'].mean()
    span = (maxd - mind).days
    
    js = f'''const {{Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,HeadingLevel,AlignmentType,WidthType,ShadingType,PageBreak,Header,Footer,PageNumber}}=require("docx");
const fs=require("fs");
const doc=new Document({{
styles:{{default:{{document:{{run:{{font:"Arial",size:24}}}}}},paragraphStyles:[
{{id:"Heading1",name:"Heading 1",basedOn:"Normal",run:{{size:36,bold:true,color:"1F4E79"}},paragraph:{{spacing:{{before:400,after:200}}}}}},
{{id:"Heading2",name:"Heading 2",basedOn:"Normal",run:{{size:28,bold:true,color:"1F4E79"}},paragraph:{{spacing:{{before:300,after:150}}}}}}]}},
sections:[{{properties:{{page:{{size:{{width:12240,height:15840}},margin:{{top:1440,right:1440,bottom:1440,left:1440}}}}}},
headers:{{default:new Header({{children:[new Paragraph({{children:[new TextRun({{text:{json.dumps(title)},size:20,color:"666666"}})]}})]}})}},
footers:{{default:new Footer({{children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun({{text:"Page ",size:20}}),new TextRun({{children:[PageNumber.CURRENT],size:20}})]}})]}})}},
children:[
new Paragraph({{heading:HeadingLevel.HEADING_1,children:[new TextRun({json.dumps(title)})]}}) ,
new Paragraph({{children:[new TextRun({{text:"Généré le {datetime.now().strftime('%d/%m/%Y')}",italics:true,color:"666666"}})]}}),
new Paragraph({{children:[]}}),
new Paragraph({{heading:HeadingLevel.HEADING_2,children:[new TextRun("Résumé")]}}),
new Paragraph({{children:[new TextRun({{text:"Tâches: ",bold:true}}),new TextRun("{n}")]}}),
new Paragraph({{children:[new TextRun({{text:"Catégories: ",bold:true}}),new TextRun("{len(cats)}")]}}),
new Paragraph({{children:[new TextRun({{text:"Durée moyenne: ",bold:true}}),new TextRun("{avg:.0f} jours")]}}),
new Paragraph({{children:[new TextRun({{text:"Période: ",bold:true}}),new TextRun("{mind.strftime('%d/%m/%Y')} → {maxd.strftime('%d/%m/%Y')} ({span}j)")]}}),
new Paragraph({{children:[]}}),
new Paragraph({{heading:HeadingLevel.HEADING_2,children:[new TextRun("Par catégorie")]}}),
new Table({{width:{{size:100,type:WidthType.PERCENTAGE}},rows:[
new TableRow({{children:[
new TableCell({{shading:{{fill:"1F4E79",type:ShadingType.CLEAR}},children:[new Paragraph({{children:[new TextRun({{text:"Catégorie",bold:true,color:"FFFFFF"}})]}})]}}),
new TableCell({{shading:{{fill:"1F4E79",type:ShadingType.CLEAR}},children:[new Paragraph({{children:[new TextRun({{text:"Tâches",bold:true,color:"FFFFFF"}})]}})]}}),
new TableCell({{shading:{{fill:"1F4E79",type:ShadingType.CLEAR}},children:[new Paragraph({{children:[new TextRun({{text:"Durée moy.",bold:true,color:"FFFFFF"}})]}})]}})]}}),
'''
    for c in cats:
        cd = data[data['categorie'] == c]
        js += f'new TableRow({{children:[new TableCell({{children:[new Paragraph({{children:[new TextRun({json.dumps(c)})]}})]}}),'
        js += f'new TableCell({{children:[new Paragraph({{children:[new TextRun("{len(cd)}")]}})]}}),'
        js += f'new TableCell({{children:[new Paragraph({{children:[new TextRun("{cd["duree_jours"].mean():.0f}j")]}})]}})]}}),\n'
    
    js += ']}}),'
    js += 'new Paragraph({children:[new PageBreak()]}),'
    
    for c in cats:
        cd = data[data['categorie'] == c]
        js += f'new Paragraph({{heading:HeadingLevel.HEADING_2,children:[new TextRun({json.dumps(c)})]}}),'
        js += f'new Paragraph({{children:[new TextRun({{text:"{len(cd)} tâches | Durée moy.: {cd["duree_jours"].mean():.0f}j",color:"666666"}})]}}),new Paragraph({{children:[]}}),'
        
        js += '''new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
new TableRow({children:[
new TableCell({shading:{fill:"E8E8E8",type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:"Tâche",bold:true})]})]}),
new TableCell({shading:{fill:"E8E8E8",type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:"Début",bold:true})]})]}),
new TableCell({shading:{fill:"E8E8E8",type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:"Fin",bold:true})]})]}),
new TableCell({shading:{fill:"E8E8E8",type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:"Durée",bold:true})]})]})]}),'''
        
        for _, r in cd.iterrows():
            js += f'new TableRow({{children:[new TableCell({{children:[new Paragraph({{children:[new TextRun({json.dumps(r["tache"])})]}})]}}),new TableCell({{children:[new Paragraph({{children:[new TextRun("{r["debut"].strftime("%d/%m/%Y")}")]}})]}}),'
            js += f'new TableCell({{children:[new Paragraph({{children:[new TextRun("{r["fin"].strftime("%d/%m/%Y")}")]}})]}}),'
            js += f'new TableCell({{children:[new Paragraph({{children:[new TextRun("{r["duree_jours"]}j")]}})]}})]}}),\n'
        
        js += ']}),new Paragraph({children:[]}),'
    
    js += ''']}]}});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("output.docx",b);console.log("OK")});'''
    return js


def run_node(script, ext):
    with tempfile.TemporaryDirectory() as tmp:
        # Créer package.json pour les dépendances locales
        pkg = {"dependencies": {"pptxgenjs": "^3.12.0", "docx": "^8.2.0"}}
        with open(os.path.join(tmp, "package.json"), "w") as f:
            json.dump(pkg, f)
        
        # Installer les dépendances npm localement
        try:
            subprocess.run(["npm", "install", "--silent"], cwd=tmp, capture_output=True, timeout=120)
        except:
            pass
        
        sf = os.path.join(tmp, "gen.js")
        of = os.path.join(tmp, f"output.{ext}")
        with open(sf, "w", encoding="utf-8") as f:
            f.write(script)
        try:
            r = subprocess.run(["node", sf], cwd=tmp, capture_output=True, text=True, timeout=60)
            if r.returncode == 0 and os.path.exists(of):
                with open(of, "rb") as f:
                    return f.read()
        except:
            pass
    return None


def main():
    st.markdown('<p class="main-header">📊 Gantt Generic</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Diagrammes de Gantt à partir de fichiers Excel/CSV</p>', unsafe_allow_html=True)
    
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("**Colonnes:** catégorie, tâche, début, fin\n\n**Dates:** YYYY-MM-DD ou DD/MM/YYYY")
        st.divider()
        tpl = pd.DataFrame({'catégorie': ['Phase 1', 'Phase 1', 'Phase 2'], 'tâche': ['Analyse', 'Design', 'Dev'], 'début': ['2025-01-01', '2025-01-15', '2025-02-01'], 'fin': ['2025-01-14', '2025-01-31', '2025-03-15']})
        buf = io.BytesIO()
        tpl.to_excel(buf, index=False, engine='openpyxl')
        st.download_button("📥 Template", buf.getvalue(), "template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    f = st.file_uploader("📁 Fichier Excel/CSV", type=['xlsx', 'xls', 'csv'])
    
    if f:
        data, err = load_data(f)
        if err:
            st.error(err)
            return
        
        st.success(f"✅ {len(data)} tâches chargées")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tâches", len(data))
        c2.metric("Catégories", data['categorie'].nunique())
        c3.metric("Durée moy.", f"{data['duree_jours'].mean():.0f}j")
        c4.metric("Période", f"{(data['fin'].max() - data['debut'].min()).days}j")
        
        with st.expander("👁️ Données"):
            st.dataframe(data, use_container_width=True, hide_index=True)
        
        st.divider()
        col1, col2 = st.columns(2)
        mode = col1.radio("Affichage", ["Global", "Par catégorie"], horizontal=True)
        title = col2.text_input("Titre", "Diagramme de Gantt")
        
        st.divider()
        if mode == "Global":
            st.markdown(f'<div class="gantt-container">{generate_svg(data, title)}</div>', unsafe_allow_html=True)
        else:
            for cat in data['categorie'].unique():
                with st.expander(cat, expanded=True):
                    cd = data[data['categorie'] == cat].reset_index(drop=True)
                    st.markdown(f'<div class="gantt-container">{generate_svg(cd, f"{title} - {cat}")}</div>', unsafe_allow_html=True)
        
        st.divider()
        st.subheader("📥 Exports")
        
        svg_cat = {c: generate_svg(data[data['categorie'] == c].reset_index(drop=True), c) for c in data['categorie'].unique()}
        
        c1, c2, c3, c4, c5 = st.columns(5)
        
        csv_buf = io.StringIO()
        exp = data.copy()
        exp['debut'] = exp['debut'].dt.strftime('%Y-%m-%d')
        exp['fin'] = exp['fin'].dt.strftime('%Y-%m-%d')
        exp.to_csv(csv_buf, index=False)
        c1.download_button("📄 CSV", csv_buf.getvalue(), f"gantt_{datetime.now():%Y%m%d}.csv", "text/csv")
        
        c2.download_button("🖼️ SVG", generate_svg(data, title), f"gantt_{datetime.now():%Y%m%d}.svg", "image/svg+xml")
        c3.download_button("🌐 HTML", generate_html(data, svg_cat, title), f"gantt_{datetime.now():%Y%m%d}.html", "text/html")
        
        if c4.button("📊 PowerPoint"):
            with st.spinner("Génération PPTX..."):
                pptx = run_node(gen_pptx(data, title), "pptx")
                if pptx:
                    st.download_button("⬇️ PPTX", pptx, f"gantt_{datetime.now():%Y%m%d}.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", key="dl_pptx")
                else:
                    st.warning("Node.js requis pour PPTX")
        
        if c5.button("📝 Word"):
            with st.spinner("Génération DOCX..."):
                docx = run_node(gen_docx(data, title), "docx")
                if docx:
                    st.download_button("⬇️ DOCX", docx, f"gantt_{datetime.now():%Y%m%d}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_docx")
                else:
                    st.warning("Node.js requis pour DOCX")
    else:
        st.info("👆 Chargez un fichier pour commencer")
        with st.expander("🎯 Démo", expanded=True):
            demo = pd.DataFrame({'categorie': ['Phase 1', 'Phase 1', 'Phase 2', 'Phase 3'], 'tache': ['Analyse', 'Design', 'Développement', 'Tests'], 'debut': [datetime(2025,1,1), datetime(2025,1,15), datetime(2025,2,1), datetime(2025,3,1)], 'fin': [datetime(2025,1,14), datetime(2025,1,31), datetime(2025,2,28), datetime(2025,3,31)]})
            demo['duree_jours'] = (demo['fin'] - demo['debut']).dt.days + 1
            st.markdown(f'<div class="gantt-container">{generate_svg(demo, "Exemple")}</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
