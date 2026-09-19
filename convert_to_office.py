# -*- coding: utf-8 -*-
import os
import sys
import re
import subprocess
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from pathlib import Path
import platform
import shutil

BASE_DIR = Path(__file__).resolve().parent
DOCX_DIR = BASE_DIR / "docx"
PDF_DIR = BASE_DIR / "PDF"
OFFICECLI_PATH = BASE_DIR.parents[1] / "00_System" / "bin" / "officecli.exe"

IS_WINDOWS = platform.system() == "Windows"
HAS_WIN32COM = False
if IS_WINDOWS:
    try:
        import win32com.client
        HAS_WIN32COM = True
    except ImportError:
        HAS_WIN32COM = False

IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')


def run_officecli_checks(document_path):
    """生成したOffice文書の構造検証と問題検出を行う。検証失敗では生成を中断しない。"""
    if not os.path.isfile(OFFICECLI_PATH):
        print(f"WARNING: OfficeCLIが見つからないため検証を省略します: {OFFICECLI_PATH}")
        return

    checks = [
        ("OpenXMLスキーマ検証", [OFFICECLI_PATH, "validate", document_path, "--json"]),
        ("文書問題の検出", [OFFICECLI_PATH, "view", document_path, "issues", "--json"]),
    ]
    for label, command in checks:
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
        if result.returncode == 0:
            print(f"OfficeCLI {label}: 完了")
            if result.stdout.strip():
                print(result.stdout.strip())
        else:
            detail = result.stderr.strip() or result.stdout.strip() or "詳細情報なし"
            print(f"WARNING: OfficeCLI {label}に失敗しました: {detail}")


def set_eastasia_font(run, font_name="游ゴシック"):
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def set_table_borders(table, color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def parse_markdown(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # フロントマターを除去
    front_matter_pattern = re.compile(r'^---\s*\n(.*?)\n---\s*\n', re.DOTALL)
    match = front_matter_pattern.match(content)
    if match:
        body = content[match.end():]
    else:
        body = content
        
    lines = body.split('\n')
    parsed = []
    idx = 0
    
    while idx < len(lines):
        raw_line = lines[idx]
        stripped = raw_line.strip()
        idx += 1
        
        if not stripped:
            continue
            
        # テンプレートタグ除外
        if stripped.startswith('{{') and stripped.endswith('}}'):
            continue

        # 水平区切り線
        if stripped in ('---', '***', '___'):
            parsed.append(('hr', ''))
            continue

        image_match = IMAGE_PATTERN.match(stripped)
        if image_match:
            alt_text, relative_path = image_match.groups()
            image_path = os.path.normpath(os.path.join(os.path.dirname(filepath), relative_path))
            parsed.append(('image', (alt_text, image_path)))
            continue
            
        # 見出し判定
        if stripped.startswith('# '):
            parsed.append(('h1', stripped[2:]))
        elif stripped.startswith('## '):
            parsed.append(('h2', stripped[3:]))
        elif stripped.startswith('### '):
            parsed.append(('h3', stripped[4:]))
        # コールアウト判定 (> [!TAG])
        elif stripped.startswith('> [!'):
            tag_match = re.match(r'^>\s*\[!(.*?)\]\s*(.*)', stripped)
            tag = tag_match.group(1).upper() if tag_match else 'NOTE'
            first_text = tag_match.group(2).strip() if tag_match and tag_match.group(2) else ''
            
            callout_lines = []
            if first_text:
                callout_lines.append(first_text)
                
            # 続く引用行を収集
            while idx < len(lines):
                next_raw = lines[idx]
                next_stripped = next_raw.strip()
                if next_stripped.startswith('>'):
                    # 引用記号を除去
                    content_line = re.sub(r'^>\s?', '', next_stripped)
                    if content_line:
                        callout_lines.append(content_line)
                    idx += 1
                else:
                    break
            
            parsed.append(('callout', (tag, " ".join(callout_lines))))
        # 通常の引用
        elif stripped.startswith('> '):
            quote_lines = [stripped[2:]]
            while idx < len(lines):
                next_raw = lines[idx]
                next_stripped = next_raw.strip()
                if next_stripped.startswith('> ') and not next_stripped.startswith('> [!'):
                    quote_lines.append(next_stripped[2:])
                    idx += 1
                else:
                    break
            parsed.append(('quote', " ".join(quote_lines)))
        # 箇条書き判定 (インデント対応)
        elif re.match(r'^\s*[\*\-]\s+', raw_line):
            indent_spaces = len(raw_line) - len(raw_line.lstrip())
            level = indent_spaces // 2
            bullet_text = re.sub(r'^\s*[\*\-]\s+', '', raw_line)
            parsed.append(('bullet', (level, bullet_text)))
        # 番号付きリスト判定 (インデント対応)
        elif re.match(r'^\s*\d+\.\s+', raw_line):
            indent_spaces = len(raw_line) - len(raw_line.lstrip())
            level = indent_spaces // 2
            num_match = re.match(r'^\s*(\d+)\.\s+(.*)', raw_line)
            if num_match:
                num_prefix = num_match.group(1)
                num_text = num_match.group(2)
                parsed.append(('num_list', (level, num_prefix, num_text)))
            else:
                parsed.append(('paragraph', stripped))
        # テーブル判定 (| で始まり | で終わる行)
        elif stripped.startswith('|') and stripped.endswith('|'):
            table_lines = [stripped]
            while idx < len(lines):
                next_raw = lines[idx]
                next_stripped = next_raw.strip()
                if next_stripped.startswith('|') and next_stripped.endswith('|'):
                    table_lines.append(next_stripped)
                    idx += 1
                else:
                    break
            
            rows = []
            for t_line in table_lines:
                cells = [c.strip() for c in t_line.split('|')[1:-1]]
                # 区切り行 (--- や :---:) を除外
                if all(c.replace('-', '').replace(':', '').strip() == '' for c in cells if c):
                    continue
                rows.append(cells)
            if rows:
                parsed.append(('table', rows))
        # 通常の段落
        else:
            parsed.append(('paragraph', stripped))
            
    return parsed


def add_formatted_runs(p, text, default_color=None, font_name="游ゴシック"):
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            bold_text = part[2:-2]
            run = p.add_run(bold_text)
            run.bold = True
        else:
            run = p.add_run(part)
            
        if default_color:
            run.font.color.rgb = default_color
        set_eastasia_font(run, font_name)
    return p


def build_docx(parsed_data, output_path):
    doc = docx.Document()
    
    # 余白設定 (上下左右 25.4mm = 1インチ)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    # スタイルのカスタム
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Yu Gothic'
    font_normal.size = Pt(10.5)
    font_normal.color.rgb = RGBColor(0x1F, 0x29, 0x33) # 濃いグレー
    
    PRIMARY_COLOR = RGBColor(0x0F, 0x76, 0x6E)     # 濃いティールグリーン
    DARK_TEAL = RGBColor(0x11, 0x5E, 0x59)
    TEXT_COLOR = RGBColor(0x1F, 0x29, 0x33)
    MUTED_COLOR = RGBColor(0x47, 0x55, 0x69)

    for item_type, content in parsed_data:
        if item_type == 'h1':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(14)
            p.paragraph_format.line_spacing = 1.2
            
            # 見出し1テキスト
            add_formatted_runs(p, content, default_color=PRIMARY_COLOR)
            for r in p.runs:
                r.font.size = Pt(18)
                r.bold = True
                
        elif item_type == 'h2':
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.keep_with_next = True
            
            add_formatted_runs(p, content, default_color=PRIMARY_COLOR)
            for r in p.runs:
                r.font.size = Pt(13.5)
                r.bold = True
                
        elif item_type == 'h3':
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            
            add_formatted_runs(p, content, default_color=DARK_TEAL)
            for r in p.runs:
                r.font.size = Pt(11.5)
                r.bold = True
                
        elif item_type == 'bullet':
            level, text = content
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            
            add_formatted_runs(p, text, default_color=TEXT_COLOR)
            
        elif item_type == 'num_list':
            level, num_prefix, text = content
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            
            # 番号プレフィックスを太字で追加
            run_num = p.add_run(f"{num_prefix}. ")
            run_num.bold = True
            run_num.font.color.rgb = PRIMARY_COLOR
            set_eastasia_font(run_num)
            
            add_formatted_runs(p, text, default_color=TEXT_COLOR)
            
        elif item_type == 'callout':
            tag, text = content
            table = doc.add_table(rows=1, cols=1)
            table.autofit = False
            table.columns[0].width = Inches(6.5)
            cell = table.cell(0, 0)
            
            # タグに応じた色分け
            tag_configs = {
                'IMPORTANT': {'fill': "FEF2F2", 'border': "DC2626", 'tag_rgb': RGBColor(0xDC, 0x26, 0x26), 'label': "【重要】"},
                'WARNING':   {'fill': "FFFBEB", 'border': "D97706", 'tag_rgb': RGBColor(0xD9, 0x77, 0x06), 'label': "【注意】"},
                'TIP':       {'fill': "F0FDF4", 'border': "16A34A", 'tag_rgb': RGBColor(0x16, 0xA3, 0x4A), 'label': "【ポイント】"},
                'NOTE':      {'fill': "F0FDFA", 'border': "0F766E", 'tag_rgb': RGBColor(0x0F, 0x76, 0x6E), 'label': "【まとめ】"},
            }
            config = tag_configs.get(tag, tag_configs['NOTE'])
            
            # セル背景色
            shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{config["fill"]}"/>')
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
            # セル枠線 (左側太線のみ)
            borders_elm = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{config["border"]}"/>'
                f'<w:top w:val="none"/>'
                f'<w:right w:val="none"/>'
                f'<w:bottom w:val="none"/>'
                f'</w:tcBorders>'
            )
            cell._tc.get_or_add_tcPr().append(borders_elm)
            
            # セル内余白 (パディング)
            mar_elm = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'<w:top w:w="120" w:type="dxa"/>'
                f'<w:bottom w:w="120" w:type="dxa"/>'
                f'<w:left w:w="180" w:type="dxa"/>'
                f'<w:right w:w="180" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            cell._tc.get_or_add_tcPr().append(mar_elm)
            
            cp = cell.paragraphs[0]
            cp.paragraph_format.space_before = Pt(2)
            cp.paragraph_format.space_after = Pt(2)
            cp.paragraph_format.line_spacing = 1.15
            
            run_tag = cp.add_run(f"{config['label']} ")
            run_tag.bold = True
            run_tag.font.color.rgb = config['tag_rgb']
            set_eastasia_font(run_tag)
            
            add_formatted_runs(cp, text, default_color=TEXT_COLOR)
            
            # テーブル下のスペース
            space_p = doc.add_paragraph()
            space_p.paragraph_format.space_before = Pt(0)
            space_p.paragraph_format.space_after = Pt(6)
            
        elif item_type == 'quote':
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            
            add_formatted_runs(p, content, default_color=MUTED_COLOR)
            
        elif item_type == 'hr':
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)
            p_border = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                f'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/>'
                f'</w:pBdr>'
            )
            p._p.get_or_add_pPr().append(p_border)
            
        elif item_type == 'image':
            alt_text, image_path = content
            if not os.path.isfile(image_path):
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                add_formatted_runs(p, f"【図版未検出】{alt_text}: {image_path}", default_color=MUTED_COLOR)
                continue
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(image_path, width=Inches(5.8))
            p.paragraph_format.space_after = Pt(3)
            if alt_text:
                caption = doc.add_paragraph()
                caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption.paragraph_format.space_after = Pt(8)
                add_formatted_runs(caption, alt_text, default_color=MUTED_COLOR)
        elif item_type == 'table':
            rows_data = content
            if not rows_data:
                continue
            col_count = max(len(r) for r in rows_data)
            table = doc.add_table(rows=len(rows_data), cols=col_count)
            table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(table, color="CBD5E1")
            
            for row_idx, row_cells in enumerate(rows_data):
                is_header = (row_idx == 0)
                row_elem = table.rows[row_idx]
                
                if is_header:
                    trPr = row_elem._tr.get_or_add_trPr()
                    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                
                for col_idx in range(col_count):
                    cell = row_elem.cells[col_idx]
                    cell_text = row_cells[col_idx] if col_idx < len(row_cells) else ""
                    
                    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
                    
                    if is_header:
                        set_cell_background(cell, "0F766E")
                    elif row_idx % 2 == 1:
                        set_cell_background(cell, "F8FAFC")
                    else:
                        set_cell_background(cell, "FFFFFF")
                    
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                    
                    # 危険度列（col_idx==2）は中央揃え、時期列（col_idx==0）も中央揃え
                    if col_idx in (0, 2):
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    lines = cell_text.replace('<br>', '\n').split('\n')
                    for l_idx, line in enumerate(lines):
                        if l_idx > 0:
                            p = cell.add_paragraph()
                            p.paragraph_format.space_before = Pt(0)
                            p.paragraph_format.space_after = Pt(0)
                            p.paragraph_format.line_spacing = 1.15
                            if col_idx in (0, 2):
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        if is_header:
                            add_formatted_runs(p, line, default_color=RGBColor(0xFF, 0xFF, 0xFF))
                            for r in p.runs:
                                r.font.size = Pt(9.5)
                                r.bold = True
                        else:
                            add_formatted_runs(p, line, default_color=TEXT_COLOR)
                            for r in p.runs:
                                r.font.size = Pt(9.0)
            
            # 列幅の調整
            if col_count == 4:
                col_widths = [Inches(1.1), Inches(1.5), Inches(0.9), Inches(2.77)]
                for row in table.rows:
                    for c_idx, w in enumerate(col_widths):
                        row.cells[c_idx].width = w
            
            p_after = doc.add_paragraph()
            p_after.paragraph_format.space_before = Pt(0)
            p_after.paragraph_format.space_after = Pt(6)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.2
            add_formatted_runs(p, content, default_color=TEXT_COLOR)
            
    doc.save(output_path)
    print(f"Successfully generated DOCX at {output_path}")


def convert_docx_to_pdf(docx_path, pdf_path):
    abs_docx = str(Path(docx_path).resolve())
    abs_pdf = str(Path(pdf_path).resolve())
    
    # 1. Windows環境かつWord COMが利用可能な場合
    if IS_WINDOWS and HAS_WIN32COM:
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(abs_docx)
                doc.SaveAs(abs_pdf, FileFormat=17)
                doc.Close()
                print(f"Successfully converted {docx_path} to {pdf_path} (via Word COM)")
                return
            finally:
                word.Quit()
        except Exception as e:
            print(f"WARNING: Word COM conversion failed: {e}. Trying LibreOffice fallback...")

    # 2. LibreOffice (soffice) によるフォールバック (macOS / Linux / Windows)
    soffice_cmd = shutil.which("soffice")
    if not soffice_cmd and platform.system() == "Darwin":
        mac_libreoffice = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if mac_libreoffice.exists():
            soffice_cmd = str(mac_libreoffice)

    if soffice_cmd:
        try:
            out_dir = str(Path(pdf_path).parent.resolve())
            cmd = [soffice_cmd, "--headless", "--convert-to", "pdf", abs_docx, "--outdir", out_dir]
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode == 0:
                print(f"Successfully converted {docx_path} to {pdf_path} (via LibreOffice)")
                return
            else:
                print(f"WARNING: LibreOffice conversion failed: {res.stderr}")
        except Exception as e:
            print(f"WARNING: LibreOffice execution failed: {e}")

    # 3. 変換エンジンが利用できない場合
    print(f"INFO: PDF変換エンジン（Word COMまたはLibreOffice）が利用できないため、PDF変換をスキップしました（DOCX生成は完了）。")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_to_office.py <markdown_file>")
        sys.exit(1)
        
    md_file = Path(sys.argv[1])
    if not md_file.exists():
        print(f"Error: {md_file} does not exist.")
        sys.exit(1)
        
    base_name = md_file.stem
    
    DOCX_DIR.mkdir(parents=True, exist_ok=True)
    PDF_DIR.mkdir(parents=True, exist_ok=True)
        
    docx_path = DOCX_DIR / f"{base_name}.docx"
    pdf_path = PDF_DIR / f"{base_name}.pdf"
    
    parsed = parse_markdown(str(md_file))
    build_docx(parsed, str(docx_path))
    run_officecli_checks(str(docx_path))
    convert_docx_to_pdf(str(docx_path), str(pdf_path))
