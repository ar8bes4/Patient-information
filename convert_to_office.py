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
import win32com.client

OFFICECLI_PATH = r"C:\Users\yert1\Documents\agy\00_System\bin\officecli.exe"
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
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.2
            add_formatted_runs(p, content, default_color=TEXT_COLOR)
            
    doc.save(output_path)
    print(f"Successfully generated DOCX at {output_path}")


def convert_docx_to_pdf(docx_path, pdf_path):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)
        
        doc = word.Documents.Open(abs_docx)
        doc.SaveAs(abs_pdf, FileFormat=17)
        doc.Close()
        print(f"Successfully converted {docx_path} to {pdf_path}")
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
        raise e
    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_to_office.py <markdown_file>")
        sys.exit(1)
        
    md_file = sys.argv[1]
    if not os.path.exists(md_file):
        print(f"Error: {md_file} does not exist.")
        sys.exit(1)
        
    base_name = os.path.splitext(os.path.basename(md_file))[0]
    
    base_dir = r"C:\Users\yert1\Documents\agy\10_Medical\Patient-information"
    docx_dir = os.path.join(base_dir, "docx")
    pdf_dir = os.path.join(base_dir, "PDF")
    
    if not os.path.exists(docx_dir):
        os.makedirs(docx_dir)
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)
        
    docx_path = os.path.join(docx_dir, f"{base_name}.docx")
    pdf_path = os.path.join(pdf_dir, f"{base_name}.pdf")
    
    parsed = parse_markdown(md_file)
    build_docx(parsed, docx_path)
    run_officecli_checks(docx_path)
    convert_docx_to_pdf(docx_path, pdf_path)
