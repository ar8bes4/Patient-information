# -*- coding: utf-8 -*-
# 患者説明文書のMarkdownをプレビュー・印刷対応のHTMLテンプレートに自動変換するスクリプト
import os
import re
import glob

# パス設定
base_dir = r"C:\Users\yert1\Documents\agy\10_Medical\Patient-information"
html_dir = os.path.join(base_dir, "HTML")

# HTML出力ディレクトリが存在しない場合は作成
if not os.path.exists(html_dir):
    os.makedirs(html_dir)

# HTMLのベーステンプレート
# サンプル.htmlの優れたデザインと表示モード（画面・高齢者印刷・紙節約印刷・スライド）の切り替えJSを内包
html_template = """<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  {meta_comments}
  <title>{title}</title>
  <style id="screen-css">
    :root {{
      --ink: #1f2933;
      --muted: #52606d;
      --line: #d9e2ec;
      --panel: #f8fafc;
      --accent: #0f766e;
      --accent-strong: #115e59;
      --warning: #9f580a;
      --warning-bg: #fff7e6;
      --urgent: #b42318;
      --urgent-bg: #fff1f0;
      --paper: #ffffff;
      --font: "Yu Gothic", "YuGothic", "Meiryo", "Segoe UI", "Arial", sans-serif;
    }}
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      margin: 0;
      color: var(--ink);
      background: #eef3f7;
      font-family: var(--font);
      font-size: 17px;
      line-height: 1.75;
    }}
    .app-toolbar {{
      position: sticky;
      top: 0;
      z-index: 10;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding: 10px 18px;
      background: rgba(255, 255, 255, 0.96);
      border-bottom: 1px solid var(--line);
      backdrop-filter: blur(8px);
    }}
    .toolbar-title {{ font-weight: 700; color: var(--accent-strong); white-space: nowrap; }}
    .toolbar-actions {{ display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; }}
    button {{
      min-height: 38px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 6px 12px;
      background: #ffffff;
      color: var(--ink);
      font: inherit;
      font-size: 14px;
      cursor: pointer;
    }}
    button[aria-pressed="true"] {{
      border-color: var(--accent);
      background: var(--accent);
      color: #ffffff;
    }}
    .slide-pagination {{ display: none; }}
    .document {{ width: min(100%, 980px); margin: 0 auto; padding: 28px 18px 56px; }}
    article {{
      background: var(--paper);
      border: 1px solid var(--line);
      box-shadow: 0 14px 35px rgba(31, 41, 51, 0.08);
    }}
    .section-block {{ padding: 28px 34px; border-bottom: 1px solid var(--line); }}
    .section-block:last-child {{ border-bottom: 0; }}
    .doc-cover {{
      padding-top: 42px;
      background: linear-gradient(180deg, #ffffff 0%, #f7fbfa 100%);
    }}
    .eyebrow {{ margin: 0 0 8px; color: var(--accent-strong); font-size: 15px; font-weight: 700; }}
    h1, h2, h3 {{ margin: 0; line-height: 1.35; }}
    h1 {{ max-width: 820px; font-size: 34px; }}
    h2 {{ margin-bottom: 14px; padding-left: 12px; border-left: 5px solid var(--accent); font-size: 24px; }}
    h3 {{ margin-bottom: 8px; color: var(--accent-strong); font-size: 18px; }}
    p, ul, ol, dl {{ margin-top: 0; margin-bottom: 16px; }}
    ul, ol {{ padding-left: 1.35em; }}
    li {{ margin: 7px 0; }}
    a {{ color: var(--accent-strong); }}
    .lead {{ max-width: 780px; margin-top: 18px; font-size: 19px; }}
    .note {{
      padding: 14px 16px;
      border-left: 5px solid var(--warning);
      background: var(--warning-bg);
      margin-bottom: 16px;
    }}
    .step-list, .risk-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 16px;
    }}
    .step-list section, .risk-grid section {{
      padding: 16px;
      border: 1px solid var(--line);
      background: var(--panel);
    }}
    .check-list {{ list-style: none; padding-left: 0; }}
    .check-list li {{ position: relative; padding-left: 30px; }}
    .check-list li::before {{
      content: "";
      position: absolute;
      left: 0;
      top: 0.58em;
      width: 16px;
      height: 16px;
      border: 2px solid var(--accent);
      border-radius: 3px;
      background: #ffffff;
    }}
    .urgent {{ border-left: 7px solid var(--urgent); background: var(--urgent-bg); }}
    .urgent h2 {{ border-left-color: var(--urgent); }}
    
    @media (max-width: 760px) {{
      body {{ font-size: 16px; }}
      .app-toolbar {{ position: static; align-items: flex-start; flex-direction: column; }}
      .toolbar-actions {{ justify-content: flex-start; }}
      .document {{ padding: 0; }}
      article {{ border: 0; box-shadow: none; }}
      .section-block {{ padding: 24px 18px; }}
      h1 {{ font-size: 28px; }}
      h2 {{ font-size: 21px; }}
      .step-list, .risk-grid {{ grid-template-columns: 1fr; }}
    }}
    @media print {{ .app-toolbar, .slide-pagination {{ display: none; }} }}
  </style>
  <style id="print-large-css" media="print">
    @page {{ size: A4; margin: 14mm 14mm 16mm; }}
    * {{ box-shadow: none !important; text-shadow: none !important; }}
    html, body {{
      margin: 0; padding: 0;
      background: #ffffff !important; color: #111111;
      font-family: var(--font);
      font-size: 14.5pt; line-height: 1.48;
    }}
    .document, article {{ width: auto; margin: 0; padding: 0; border: 0; background: #ffffff; }}
    .section-block {{ break-inside: avoid; page-break-inside: avoid; padding: 0 0 7mm; border: 0; }}
    .doc-cover {{ padding-top: 0; background: #ffffff; }}
    .eyebrow {{ margin: 0 0 3mm; font-size: 12pt; font-weight: 700; }}
    h1, h2, h3 {{ break-after: avoid; page-break-after: avoid; color: #111111; line-height: 1.28; }}
    h1 {{ margin: 0 0 6mm; font-size: 24pt; }}
    h2 {{ margin: 0 0 4mm; padding: 0 0 1.5mm; border: 0; border-bottom: 1.5pt solid #111111; font-size: 18pt; }}
    h3 {{ margin: 0 0 2mm; font-size: 15pt; }}
    p, ul, ol, dl {{ margin-top: 0; margin-bottom: 4mm; }}
    ul, ol {{ padding-left: 1.35em; }}
    li {{ margin: 1.5mm 0; }}
    a {{ color: #111111; text-decoration: none; }}
    .lead {{ font-size: 15.5pt; }}
    .note, .urgent {{
      break-inside: avoid; page-break-inside: avoid; padding: 3mm;
      border: 1.2pt solid #333333; background: #ffffff !important;
    }}
    .step-list, .risk-grid {{ display: block; }}
    .step-list section, .risk-grid section {{
      break-inside: avoid; page-break-inside: avoid; margin: 0 0 3mm; padding: 3mm;
      border: 1pt solid #444444; background: #ffffff !important;
    }}
    .check-list {{ list-style: square; padding-left: 1.35em; }}
    .check-list li {{ padding-left: 0; }}
    .check-list li::before {{ display: none; }}
  </style>
  <style id="print-compact-css" media="not all">
    @page {{ size: A4; margin: 10mm 10mm 11mm; }}
    * {{ box-shadow: none !important; text-shadow: none !important; }}
    html, body {{
      margin: 0; padding: 0;
      background: #ffffff !important; color: #111111;
      font-family: var(--font);
      font-size: 10.8pt; line-height: 1.28;
    }}
    .document, article {{ width: auto; margin: 0; padding: 0; border: 0; background: #ffffff; }}
    .section-block {{ break-inside: avoid; page-break-inside: avoid; padding: 0 0 4mm; border: 0; }}
    .doc-cover {{ padding-top: 0; background: #ffffff; }}
    .eyebrow {{ margin: 0 0 1mm; font-size: 9.5pt; font-weight: 700; }}
    h1, h2, h3 {{ break-after: avoid; page-break-after: avoid; color: #111111; line-height: 1.18; }}
    h1 {{ margin: 0 0 3mm; font-size: 18pt; }}
    h2 {{ margin: 0 0 2mm; padding: 0 0 1mm; border: 0; border-bottom: 1pt solid #111111; font-size: 13.5pt; }}
    h3 {{ margin: 0 0 1mm; font-size: 11.5pt; }}
    p, ul, ol, dl {{ margin-top: 0; margin-bottom: 2.5mm; }}
    ul, ol {{ padding-left: 1.2em; }}
    li {{ margin: 0.8mm 0; }}
    a {{ color: #111111; text-decoration: none; }}
    .lead {{ font-size: 11.5pt; }}
    .note, .urgent {{
      break-inside: avoid; page-break-inside: avoid; padding: 2mm;
      border: 0.8pt solid #333333; background: #ffffff !important;
    }}
    .step-list, .risk-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 2mm; }}
    .step-list section, .risk-grid section {{
      break-inside: avoid; page-break-inside: avoid; padding: 2mm;
      border: 0.8pt solid #444444; background: #ffffff !important;
    }}
    .check-list {{ list-style: square; padding-left: 1.2em; }}
    .check-list li {{ padding-left: 0; }}
    .check-list li::before {{ display: none; }}
  </style>
  <style id="slides-css" media="not all">
    body[data-mode="slides"] {{ overflow: hidden; background: #0f172a; }}
    body[data-mode="slides"] .app-toolbar {{
      background: rgba(15, 23, 42, 0.94); border-bottom-color: rgba(255, 255, 255, 0.16);
    }}
    body[data-mode="slides"] .toolbar-title {{ color: #ccfbf1; }}
    body[data-mode="slides"] button {{
      border-color: rgba(255, 255, 255, 0.22); background: rgba(255, 255, 255, 0.08); color: #f8fafc;
    }}
    body[data-mode="slides"] button[aria-pressed="true"] {{ background: #14b8a6; color: #062926; }}
    body[data-mode="slides"] .document {{
      width: 100vw; height: calc(100vh - 59px); margin: 0; padding: 0;
      overflow-x: auto; overflow-y: hidden; scroll-snap-type: x mandatory;
    }}
    body[data-mode="slides"] article {{
      display: flex; width: max-content; height: 100%; border: 0; box-shadow: none; background: transparent;
    }}
    body[data-mode="slides"] .section-block {{
      display: flex; flex-direction: column; justify-content: center;
      width: 100vw; height: 100%; padding: clamp(28px, 6vw, 76px);
      overflow-y: auto; scroll-snap-align: start; border: 0; background: #f8fafc;
    }}
    body[data-mode="slides"] .doc-cover {{ background: #ecfeff; }}
    body[data-mode="slides"] h1 {{ max-width: 1000px; font-size: clamp(40px, 6vw, 78px); }}
    body[data-mode="slides"] h2 {{
      margin-bottom: 28px; padding-left: 18px; border-left-width: 8px; font-size: clamp(30px, 4vw, 54px);
    }}
    body[data-mode="slides"] h3 {{ font-size: clamp(20px, 2vw, 28px); }}
    body[data-mode="slides"] p, body[data-mode="slides"] li {{
      max-width: 1100px; font-size: clamp(20px, 2vw, 30px); line-height: 1.55;
    }}
    body[data-mode="slides"] .lead {{ font-size: clamp(24px, 2.5vw, 34px); }}
    body[data-mode="slides"] .step-list, body[data-mode="slides"] .risk-grid {{
      grid-template-columns: repeat(2, minmax(320px, 1fr)); gap: 18px;
    }}
    body[data-mode="slides"] .note, body[data-mode="slides"] .urgent,
    body[data-mode="slides"] .step-list section, body[data-mode="slides"] .risk-grid section {{
      background: #ffffff; border-color: #cbd5e1;
    }}
    .slide-pagination {{ display: none; }}
    body[data-mode="slides"] .slide-pagination {{
      position: fixed; right: 22px; bottom: 18px; z-index: 20;
      display: flex; align-items: center; gap: 12px; min-width: 148px; justify-content: center;
      padding: 8px 10px; border: 1px solid rgba(255, 255, 255, 0.22); border-radius: 8px;
      background: rgba(15, 23, 42, 0.88); color: #f8fafc; font-size: 16px; line-height: 1;
    }}
    body[data-mode="slides"] .slide-nav {{
      display: inline-flex; align-items: center; justify-content: center;
      width: 34px; min-height: 34px; padding: 0;
      border-color: rgba(255, 255, 255, 0.24); border-radius: 6px;
      background: rgba(255, 255, 255, 0.1); color: #ffffff; font-size: 24px; line-height: 1;
    }}
    @media print {{ .slide-pagination {{ display: none !important; }} }}
  </style>
</head>
<body>
  <header class="app-toolbar" aria-label="表示切り替え">
    <div class="toolbar-title">{title}</div>
    <div class="toolbar-actions">
      <button type="button" data-mode="screen" aria-pressed="true">画面</button>
      <button type="button" data-mode="large" aria-pressed="false">高齢者印刷</button>
      <button type="button" data-mode="compact" aria-pressed="false">紙節約印刷</button>
      <button type="button" data-mode="slides" aria-pressed="false">スライド</button>
      <button type="button" data-action="print">印刷/PDF</button>
    </div>
  </header>

  <main class="document" id="top">
    <article>
{content}
    </article>
  </main>

  <div class="slide-pagination" aria-live="polite" aria-label="スライド番号">
    <button type="button" class="slide-nav" data-slide-nav="prev" aria-label="前のスライド">‹</button>
    <span><span id="current-slide">1</span> / <span id="total-slides">1</span></span>
    <button type="button" class="slide-nav" data-slide-nav="next" aria-label="次のスライド">›</button>
  </div>

  <script>
    const printLarge = document.getElementById("print-large-css");
    const printCompact = document.getElementById("print-compact-css");
    const slides = document.getElementById("slides-css");
    const buttons = document.querySelectorAll("[data-mode]");
    const documentFrame = document.querySelector(".document");
    const slideBlocks = Array.from(document.querySelectorAll(".section-block"));
    const currentSlide = document.getElementById("current-slide");
    const totalSlides = document.getElementById("total-slides");
    let activeSlideIndex = 0;

    function setMode(mode) {{
      document.body.dataset.mode = mode;
      printLarge.media = mode === "large" || mode === "screen" || mode === "slides" ? "print" : "not all";
      printCompact.media = mode === "compact" ? "print" : "not all";
      slides.media = mode === "slides" ? "screen" : "not all";
      buttons.forEach((button) => {{
        button.setAttribute("aria-pressed", String(button.dataset.mode === mode));
      }});
      if (mode === "slides") {{
        goToSlide(activeSlideIndex);
      }}
    }}

    function updatePagination(index) {{
      activeSlideIndex = Math.max(0, Math.min(index, slideBlocks.length - 1));
      currentSlide.textContent = String(activeSlideIndex + 1);
      totalSlides.textContent = String(slideBlocks.length);
    }}

    function goToSlide(index) {{
      updatePagination(index);
      documentFrame.scrollTo({{
        left: activeSlideIndex * documentFrame.clientWidth,
        behavior: "smooth"
      }});
    }}

    function syncPaginationFromScroll() {{
      if (document.body.dataset.mode !== "slides") return;
      const width = Math.max(1, documentFrame.clientWidth);
      updatePagination(Math.round(documentFrame.scrollLeft / width));
    }}

    buttons.forEach((button) => {{
      button.addEventListener("click", () => setMode(button.dataset.mode));
    }});

    document.querySelector("[data-slide-nav='prev']").addEventListener("click", () => {{
      goToSlide(activeSlideIndex - 1);
    }});

    document.querySelector("[data-slide-nav='next']").addEventListener("click", () => {{
      goToSlide(activeSlideIndex + 1);
    }});

    document.addEventListener("keydown", (event) => {{
      if (document.body.dataset.mode !== "slides") return;
      if (event.key === "ArrowRight" || event.key === "PageDown" || event.key === " ") {{
        event.preventDefault();
        goToSlide(activeSlideIndex + 1);
      }}
      if (event.key === "ArrowLeft" || event.key === "PageUp") {{
        event.preventDefault();
        goToSlide(activeSlideIndex - 1);
      }}
      if (event.key === "Home") {{
        event.preventDefault();
        goToSlide(0);
      }}
      if (event.key === "End") {{
        event.preventDefault();
        goToSlide(slideBlocks.length - 1);
      }}
    }});

    documentFrame.addEventListener("scroll", syncPaginationFromScroll);
    window.addEventListener("resize", () => {{
      if (document.body.dataset.mode === "slides") {{
        goToSlide(activeSlideIndex);
      }}
    }});

    document.querySelector("[data-action='print']").addEventListener("click", () => {{
      window.print();
    }});

    updatePagination(0);
    setMode("screen");
  </script>
</body>
</html>
"""

# Markdownのパースとセクション分割ロジック
def parse_markdown_to_sections(md_text):
    import datetime
    lines = md_text.splitlines()
    
    # 1. フロントマター (YAML) の抽出
    title = "患者説明資料"
    description = ""
    start_idx = 0
    metadata = {
        "reviewed_by": "",
        "review_date": "",
        "evidence_source": "",
        "change_reason": "",
        "version": ""
    }
    
    if len(lines) > 0 and lines[0].strip() == "---":
        fm_lines = []
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                start_idx = i + 1
                break
            fm_lines.append(lines[i])
        
        # フロントマターの解析
        for line in fm_lines:
            m_title = re.match(r"^title:\s*(.*)$", line)
            if m_title:
                title = m_title.group(1).strip().strip('"').strip("'")
                continue
            m_desc = re.match(r"^description:\s*(.*)$", line)
            if m_desc:
                description = m_desc.group(1).strip().strip('"').strip("'")
                continue
            
            # その他の追跡メタデータを抽出
            for key in metadata.keys():
                m_meta = re.match(rf"^{key}:\s*(.*)$", line)
                if m_meta:
                    metadata[key] = m_meta.group(1).strip().strip('"').strip("'")
                    break

    # 画像リストの抽出
    image_paths = re.findall(r"!\[.*?\]\((.*?)\)", md_text)
    metadata["images"] = image_paths
    metadata["build_datetime"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    body_lines = lines[start_idx:]
    
    # 2. 最初にあるH1 (# タイトル) を抽出し、フロントマターにタイトルがない場合はそれをタイトルにする
    content_lines = []
    for line in body_lines:
        if (line.startswith("# ") or line.startswith("# **")) and title == "患者説明資料":
            # タイトル行を抽出
            raw_title = line.replace("#", "").replace("**", "").strip()
            if raw_title:
                title = raw_title
            continue
        content_lines.append(line)
        
    # 3. 行ごとに解析してHTMLセクションに分類
    sections = []
    current_section_content = []
    current_section_title = ""
    is_cover = True # 最初のH2に到達するまではカバーセクション
    
    # リスト状態管理
    in_ul = False
    in_ol = False
    in_p = False
    in_note = False
    in_urgent = False
    
    def close_all_tags(content_list):
        nonlocal in_ul, in_ol, in_p, in_note, in_urgent
        out = []
        if in_ul:
            out.append("        </ul>")
            in_ul = False
        if in_ol:
            out.append("        </ol>")
            in_ol = False
        if in_p:
            out.append("        </p>")
            in_p = False
        if in_note or in_urgent:
            out.append("        </div>")
            in_note = False
            in_urgent = False
        if out:
            content_list.extend(out)

    for line in content_lines:
        cleaned_line = line.strip()
        
        # H2見出し (セクションの境界)
        if cleaned_line.startswith("##") and not cleaned_line.startswith("###"):
            h2_text = re.sub(r"^##\s*", "", cleaned_line).replace("**", "").strip()
            # 空の H2 (## のみ等) は無視して完全に除外する
            if not h2_text:
                continue
                
            # 既存のセクションをクローズして保存
            close_all_tags(current_section_content)
            
            if is_cover:
                # カバーセクションを保存
                sections.append({
                    "is_cover": True,
                    "title": title,
                    "content": "\n".join(current_section_content)
                })
                is_cover = False
            else:
                # 通常セクションを保存
                sections.append({
                    "is_cover": False,
                    "title": current_section_title,
                    "content": "\n".join(current_section_content)
                })
            
            # 新しいセクションをスタート
            current_section_content = []
            current_section_title = h2_text
            current_section_content.append(f"      <h2>{h2_text}</h2>")
            continue
            
        # 空行
        if not cleaned_line:
            if in_p:
                current_section_content.append("        </p>")
                in_p = False
            continue

        # リストマークアップのクローズ判定
        is_list_line = (cleaned_line.startswith("- ") or cleaned_line.startswith("* ") or 
                        cleaned_line.startswith("+ ") or re.match(r"^\d+\.\s+", cleaned_line))
        
        if not is_list_line:
            if in_ul:
                current_section_content.append("        </ul>")
                in_ul = False
            if in_ol:
                current_section_content.append("        </ol>")
                in_ol = False
                
        # 箇条書きリスト (UL)
        if cleaned_line.startswith("- ") or cleaned_line.startswith("* ") or cleaned_line.startswith("+ "):
            if not in_ul:
                # Pがオープンしていれば閉じる
                if in_p:
                    current_section_content.append("        </p>")
                    in_p = False
                current_section_content.append("        <ul>")
                in_ul = True
            list_text = re.sub(r"^[-*+]\s+", "", cleaned_line)
            list_text = format_inline_elements(list_text)
            current_section_content.append(f"          <li>{list_text}</li>")
            continue
            
        # 番号付きリスト (OL)
        ol_match = re.match(r"^(\d+)\.\s+(.*)$", cleaned_line)
        if ol_match:
            if not in_ol:
                if in_p:
                    current_section_content.append("        </p>")
                    in_p = False
                current_section_content.append("        <ol>")
                in_ol = True
            list_text = format_inline_elements(ol_match.group(2))
            current_section_content.append(f"          <li>{list_text}</li>")
            continue

        # H3中見出し
        if cleaned_line.startswith("###"):
            h3_text = re.sub(r"^###\s*", "", cleaned_line).replace("**", "").strip()
            if not h3_text:
                continue
            close_all_tags(current_section_content)
            current_section_content.append(f"        <h3>{h3_text}</h3>")
            continue

        # 注記 / コールアウト (引出し線・引用)
        if cleaned_line.startswith("> "):
            quote_text = cleaned_line.replace("> ", "").strip()
            
            # 特別な警告コールアウトか？
            is_warning = quote_text.startswith("[!NOTE]") or quote_text.startswith("[!TIP]") or quote_text.startswith("[!IMPORTANT]")
            is_urgent = quote_text.startswith("[!WARNING]") or quote_text.startswith("[!CAUTION]")
            
            if is_warning:
                close_all_tags(current_section_content)
                current_section_content.append('        <div class="note">')
                in_note = True
                continue
            elif is_urgent:
                close_all_tags(current_section_content)
                current_section_content.append('        <div class="note urgent">')
                in_urgent = True
                continue
            
            # 通常の引用行
            if not (in_note or in_urgent):
                # デフォルトで注記パネルに入れる
                close_all_tags(current_section_content)
                current_section_content.append('        <div class="note">')
                in_note = True
            
            quote_text = format_inline_elements(quote_text)
            current_section_content.append(f"          <p>{quote_text}</p>")
            continue
            
        # 引用ブロックが途切れたら閉じる
        if not cleaned_line.startswith("> ") and (in_note or in_urgent):
            current_section_content.append("        </div>")
            in_note = False
            in_urgent = False

        # 通常の文章（段落）
        line_text = format_inline_elements(cleaned_line)
        if not in_p:
            current_section_content.append("        <p>")
            in_p = True
        current_section_content.append(f"          {line_text}")

    # 残ったタグを閉じる
    close_all_tags(current_section_content)
    
    # 最後のセクションを保存
    if is_cover:
        sections.append({
            "is_cover": True,
            "title": title,
            "content": "\n".join(current_section_content)
        })
    else:
        sections.append({
            "is_cover": False,
            "title": current_section_title,
            "content": "\n".join(current_section_content)
        })
        
    return title, description, sections, metadata

# インライン要素（太字、斜体、改行）のフォーマット
def format_inline_elements(text):
    # 太字: **text** または __text__ -> <strong>
    text = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"__([^_]+)__", r"<strong>\1</strong>", text)
    # 改行記号: `  ` (スペース2つ) または `~` または `<br>` -> <br>
    text = text.replace("  ", "<br>")
    text = text.replace("~", "<br>")
    return text

# セクションのHTMLブロックを生成
def generate_sections_html(title, description, sections):
    html_blocks = []
    
    for sec in sections:
        if sec["is_cover"]:
            # カバーセクションのマークアップ
            lead_html = f'        <p class="lead">{description}</p>' if description else ""
            block = f"""      <section class="doc-cover section-block">
        <p class="eyebrow">患者さん・ご家族への説明資料</p>
        <h1>{title}</h1>
{lead_html}
{sec['content']}
      </section>"""
        else:
            # 通常セクションのマークアップ
            block = f"""      <section class="section-block">
{sec['content']}
      </section>"""
        html_blocks.append(block)
        
    return "\n\n".join(html_blocks)

# メイン処理：Markdownファイルの一括変換
def convert_all_markdowns():
    md_files = glob.glob(os.path.join(base_dir, "src", "*.md"))
    success_count = 0
    error_files = []
    
    print(f"Total markdown files found: {len(md_files)}")
    
    for filepath in md_files:
        filename = os.path.basename(filepath)
        # 索引ファイルやrules.md等は変換対象から除外
        if filename in ["README.md", "CHANGELOG.md", "rules.md", "Gitプロジェクト一覧.md", "Patient-information.md"]:
            continue
            
        print(f"Converting {filename}...")
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                md_text = f.read()
                
            title, description, sections, metadata = parse_markdown_to_sections(md_text)
            content_html = generate_sections_html(title, description, sections)
            
            # メタデータコメント文字列の生成
            meta_comments_lines = [
                "<!--",
                f"  Meta:reviewed_by: {metadata.get('reviewed_by', '')}",
                f"  Meta:review_date: {metadata.get('review_date', '')}",
                f"  Meta:evidence_source: {metadata.get('evidence_source', '')}",
                f"  Meta:change_reason: {metadata.get('change_reason', '')}",
                f"  Meta:version: {metadata.get('version', '')}",
                f"  Meta:build_datetime: {metadata.get('build_datetime', '')}",
                f"  Meta:images: {', '.join(metadata.get('images', []))}",
                "-->"
            ]
            meta_comments_str = "\n".join(meta_comments_lines)
            
            # HTMLベーステンプレートに流し込む
            final_html = html_template.format(
                title=title, 
                content=content_html, 
                meta_comments=meta_comments_str
            )
            
            # 出力ファイル名を作成 (.md から .html)
            out_filename = os.path.splitext(filename)[0] + ".html"
            out_filepath = os.path.join(html_dir, out_filename)
            
            with open(out_filepath, 'w', encoding='utf-8') as out_f:
                out_f.write(final_html)
                
            success_count += 1
        except Exception as e:
            print(f"ERROR converting {filename}: {e}")
            error_files.append((filename, str(e)))
            
    print("\n=== Conversion Result ===")
    print(f"Successfully converted: {success_count} files.")
    if error_files:
        print(f"Failed to convert: {len(error_files)} files.")
        for name, err in error_files:
            print(f" - {name}: {err}")

if __name__ == "__main__":
    convert_all_markdowns()
