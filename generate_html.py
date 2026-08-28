# -*- coding: utf-8 -*-
"""患者説明文書のMarkdownから、試作・閲覧用HTMLを生成する。"""
import argparse
import html
import os
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "HTML")
IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
FRONT_MATTER_PATTERN = re.compile(r'^---\s*\n(.*?)\n---\s*\n', re.DOTALL)


def front_matter_and_body(text):
    match = FRONT_MATTER_PATTERN.match(text)
    if not match:
        return {}, text
    metadata = {}
    for line in match.group(1).splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            metadata[key.strip()] = value.strip().strip('"').strip("'")
    return metadata, text[match.end():]


def inline(text):
    escaped = html.escape(text.strip())
    escaped = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', escaped)
    return escaped


def parse_blocks(body, source_dir):
    blocks, lines, index = [], body.splitlines(), 0
    while index < len(lines):
        raw = lines[index]
        text = raw.strip()
        index += 1
        if not text or (text.startswith("{{") and text.endswith("}}")):
            continue
        if text in {"---", "***", "___"}:
            blocks.append(("hr", ""))
            continue
        image_match = IMAGE_PATTERN.match(text)
        if image_match:
            alt, relative_path = image_match.groups()
            path = os.path.normpath(os.path.join(source_dir, relative_path))
            blocks.append(("image", (alt, path)))
            continue
        if text.startswith("# "):
            blocks.append(("h1", text[2:]))
        elif text.startswith("## "):
            blocks.append(("h2", text[3:]))
        elif text.startswith("### "):
            blocks.append(("h3", text[4:]))
        elif text.startswith("> [!"):
            marker = text[3:].split("]", 1)[0].upper()
            message = []
            while index < len(lines) and lines[index].strip().startswith(">"):
                message.append(lines[index].strip()[1:].strip())
                index += 1
            blocks.append(("callout", (marker, " ".join(message))))
        elif text.startswith("> "):
            quote = text[2:]
            if quote.startswith("{{") and quote.endswith("}}"):
                continue
            blocks.append(("quote", quote))
        elif text.startswith(("- ", "* ")):
            items = [text[2:]]
            while index < len(lines) and lines[index].strip().startswith(("- ", "* ")):
                items.append(lines[index].strip()[2:])
                index += 1
            blocks.append(("list", items))
        elif re.match(r"^\d+\.\s+", text):
            items = [re.sub(r"^\d+\.\s+", "", text)]
            while index < len(lines) and re.match(r"^\d+\.\s+", lines[index].strip()):
                items.append(re.sub(r"^\d+\.\s+", "", lines[index].strip()))
                index += 1
            blocks.append(("ordered_list", items))
        else:
            blocks.append(("p", raw.strip()))
    return blocks


def render_blocks(blocks):
    rendered = []
    for kind, content in blocks:
        if kind in {"h1", "h2", "h3", "p", "quote"}:
            tag = {"h1": "h1", "h2": "h2", "h3": "h3", "p": "p", "quote": "blockquote"}[kind]
            rendered.append(f"<{tag}>{inline(content)}</{tag}>")
        elif kind == "list":
            rendered.append("<ul>" + "".join(f"<li>{inline(item)}</li>" for item in content) + "</ul>")
        elif kind == "ordered_list":
            rendered.append("<ol>" + "".join(f"<li>{inline(item)}</li>" for item in content) + "</ol>")
        elif kind == "hr":
            rendered.append("<hr>")
        elif kind == "callout":
            marker, message = content
            level = "important" if marker in {"IMPORTANT", "WARNING", "CAUTION"} else "note"
            rendered.append(f'<aside class="callout {level}"><p class="callout-label">{html.escape(marker)}</p><p>{inline(message)}</p></aside>')
        elif kind == "image":
            alt, image_path = content
            if os.path.isfile(image_path):
                relative_path = os.path.relpath(image_path, OUTPUT_DIR).replace("\\", "/")
                rendered.append(f'<figure><img src="{html.escape(relative_path)}" alt="{html.escape(alt)}"><figcaption>{inline(alt)}</figcaption></figure>')
            else:
                rendered.append(f'<aside class="callout note"><p class="callout-label">図版未検出</p><p>{inline(alt)}</p></aside>')
    return "\n".join(rendered)


def build_html(title, status, body_html, source_name):
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    safe_title = html.escape(title)
    return f'''<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="source" content="{html.escape(source_name)}">
  <meta name="generated" content="{generated_at}">
  <title>{safe_title}</title>
  <style>
    :root {{ --teal:#0f766e; --teal-dark:#115e59; --ink:#1f2933; --muted:#52606d; --line:#d9e2ec; --paper:#fff; --panel:#f6fbfa; --urgent:#b42318; --urgent-bg:#fff1f0; }}
    * {{ box-sizing:border-box; }} body {{ margin:0; color:var(--ink); background:#eef3f7; font-family:"Yu Gothic","YuGothic","Meiryo",sans-serif; font-size:17px; line-height:1.8; }}
    .bar {{ padding:12px 20px; color:#fff; background:var(--teal-dark); font-size:14px; }} .document {{ width:min(100%,960px); margin:28px auto 64px; background:var(--paper); box-shadow:0 14px 35px rgba(31,41,51,.10); }}
    header {{ padding:44px 52px 32px; border-bottom:1px solid var(--line); background:linear-gradient(135deg,#fff 0%,#edf8f6 100%); }} .eyebrow {{ margin:0 0 10px; color:var(--teal-dark); font-size:14px; font-weight:700; letter-spacing:.05em; }} h1 {{ margin:0; color:var(--ink); font-size:clamp(28px,4vw,42px); line-height:1.35; }}
    main {{ padding:12px 52px 42px; }} h2 {{ margin:42px 0 16px; padding-left:14px; border-left:6px solid var(--teal); color:var(--teal-dark); font-size:25px; line-height:1.4; }} h3 {{ margin:28px 0 10px; color:var(--teal-dark); font-size:19px; }} p {{ margin:0 0 16px; }} ul, ol {{ margin:0 0 20px; padding-left:1.45em; }} li {{ margin:6px 0; }} hr {{ margin:28px 0; border:0; border-top:1px solid var(--line); }} strong {{ color:#153e3b; }} blockquote {{ margin:20px 0; padding:12px 18px; border-left:4px solid #93c5bd; background:var(--panel); color:var(--muted); }}
    .callout {{ margin:24px 0; padding:16px 18px; border-left:6px solid var(--teal); background:var(--panel); }} .callout.important {{ border-color:var(--urgent); background:var(--urgent-bg); }} .callout p {{ margin:0; }} .callout-label {{ margin-bottom:5px !important; color:var(--teal-dark); font-size:13px; font-weight:700; letter-spacing:.06em; }} .important .callout-label {{ color:var(--urgent); }} figure {{ margin:28px auto; text-align:center; }} figure img {{ max-width:100%; max-height:520px; border:1px solid var(--line); border-radius:10px; background:#fff; }} figcaption {{ margin-top:8px; color:var(--muted); font-size:14px; }} footer {{ padding:18px 52px; border-top:1px solid var(--line); color:var(--muted); font-size:13px; }}
    @media print {{ body {{ background:#fff; }} .bar {{ display:none; }} .document {{ width:100%; margin:0; box-shadow:none; }} }} @media (max-width:640px) {{ .document {{ margin:0; }} header,main,footer {{ padding-left:22px; padding-right:22px; }} body {{ font-size:16px; }} }}
  </style>
</head>
<body>
  <div class="bar">患者説明文書　|　試作・閲覧版　|　医学的内容は医師が確認します</div>
  <article class="document">
    <header><p class="eyebrow">PATIENT INFORMATION　{html.escape(status or 'draft')}</p><h1>{safe_title}</h1></header>
    <main>{body_html}</main>
    <footer>正本: {html.escape(source_name)}　／　生成日時: {generated_at}</footer>
  </article>
</body>
</html>'''


def generate(source_path):
    source_path = os.path.abspath(source_path)
    with open(source_path, encoding="utf-8") as file:
        metadata, body = front_matter_and_body(file.read())
    blocks = parse_blocks(body, os.path.dirname(source_path))
    title = metadata.get("title") or next((content for kind, content in blocks if kind == "h1"), os.path.splitext(os.path.basename(source_path))[0])
    if blocks and blocks[0][0] == "h1":
        blocks = blocks[1:]
    output = build_html(title, metadata.get("status", "draft"), render_blocks(blocks), os.path.basename(source_path))
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, os.path.splitext(os.path.basename(source_path))[0] + ".html")
    with open(output_path, "w", encoding="utf-8", newline="\n") as file:
        file.write(output)
    print(f"HTMLを生成しました: {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="患者説明文書の試作・閲覧用HTMLを生成します。")
    parser.add_argument("markdown_file", help="src配下のMarkdownファイル")
    args = parser.parse_args()
    if not os.path.isfile(args.markdown_file):
        parser.error(f"ファイルが見つかりません: {args.markdown_file}")
    generate(args.markdown_file)
