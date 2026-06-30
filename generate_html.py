# -*- coding: utf-8 -*-
# 患者説明文書のMarkdownをプレビュー・印刷対応のHTMLテンプレートに自動変換するスクリプト
import os
import re
import argparse
import html

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
    .print-consent {{ display: none; }}
    .section-layout {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) 220px;
      gap: 24px;
      align-items: start;
    }}
    .section-text {{ min-width: 0; }}
    .section-visual {{
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 180px;
      border: 1px solid #cfe5e2;
      background: #f7fbfa;
      padding: 16px;
    }}
    .section-visual svg {{
      width: 100%;
      max-width: 180px;
      height: auto;
    }}
    .section-visual figure {{
      width: 100%;
      margin: 0;
    }}
    .section-visual img {{
      display: block;
      width: 100%;
      max-width: 320px;
      height: auto;
      margin: 0 auto;
    }}
    .section-visual figcaption {{
      margin-top: 10px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
      text-align: center;
    }}
    .slide-summary {{ display: none; }}
    
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
      .section-layout {{ grid-template-columns: 1fr; }}
      .section-visual {{ min-height: 140px; }}
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
      font-size: 16pt; line-height: 1.55;
    }}
    .document, article {{ width: auto; margin: 0; padding: 0; border: 0; background: #ffffff; }}
    .print-consent {{ display: block; break-after: page; page-break-after: always; padding: 0 0 7mm; }}
    .print-consent h1 {{ margin: 0 0 5mm; font-size: 24pt; text-align: center; }}
    .consent-meta-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 3mm; margin: 0 0 5mm; }}
    .consent-field {{ min-height: 12mm; padding: 2.5mm; border: 1pt solid #333333; }}
    .consent-label {{ display: block; margin-bottom: 1mm; font-size: 10.5pt; font-weight: 700; }}
    .consent-summary {{ margin: 0 0 5mm; padding: 3mm; border: 1pt solid #333333; }}
    .consent-summary h2 {{ margin: 0 0 2mm; padding: 0; border: 0; font-size: 15pt; }}
    .consent-summary ul {{ margin-bottom: 0; }}
    .signature-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 5mm; margin-top: 5mm; }}
    .signature-box {{ min-height: 28mm; padding: 3mm; border: 1pt solid #333333; }}
    .signature-line {{ margin-top: 8mm; border-bottom: 1pt solid #333333; height: 8mm; }}
    .print-note {{ margin-top: 4mm; font-size: 10.5pt; }}
    
    /* 改ページ制御の最適化: セクション全体のbreak-insideを解除し、子要素と見出し泣き別れを制御 */
    .section-block {{ break-inside: auto; page-break-inside: auto; padding: 0 0 7mm; border: 0; }}
    p, li, .note, .urgent {{ break-inside: avoid; page-break-inside: avoid; }}
    
    .doc-cover {{ padding-top: 0; background: #ffffff; }}
    .eyebrow {{ margin: 0 0 3mm; font-size: 12pt; font-weight: 700; }}
    h1, h2, h3 {{ break-after: avoid; page-break-after: avoid; color: #111111; line-height: 1.28; }}
    h1 {{ margin: 0 0 6mm; font-size: 27pt; }}
    h2 {{ margin: 0 0 4mm; padding: 0 0 1.5mm; border: 0; border-bottom: 1.5pt solid #111111; font-size: 20pt; }}
    h3 {{ margin: 0 0 2mm; font-size: 16pt; }}
    p, ul, ol, dl {{ margin-top: 0; margin-bottom: 4mm; }}
    ul, ol {{ padding-left: 1.35em; }}
    li {{ margin: 1.5mm 0; }}
    a {{ color: #111111; text-decoration: none; }}
    .lead {{ font-size: 17pt; }}
    .section-layout {{ display: grid; grid-template-columns: minmax(0, 1fr) 38mm; gap: 7mm; }}
    .section-visual {{ min-height: 32mm; padding: 2mm; border: 0.8pt solid #999999; background: #ffffff !important; }}
    .section-visual svg {{ max-width: 34mm; }}
    .section-visual img {{ max-width: 42mm; }}
    .section-visual figcaption {{ font-size: 8.5pt; }}
    .note, .urgent {{
      padding: 3mm;
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
    .print-consent {{ display: block; break-after: page; page-break-after: always; padding: 0 0 4mm; }}
    .print-consent h1 {{ margin: 0 0 3mm; font-size: 17pt; text-align: center; }}
    .consent-meta-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 2mm; margin: 0 0 3mm; }}
    .consent-field {{ min-height: 9mm; padding: 1.8mm; border: 0.8pt solid #333333; }}
    .consent-label {{ display: block; margin-bottom: 0.8mm; font-size: 8.5pt; font-weight: 700; }}
    .consent-summary {{ margin: 0 0 3mm; padding: 2mm; border: 0.8pt solid #333333; }}
    .consent-summary h2 {{ margin: 0 0 1mm; padding: 0; border: 0; font-size: 11.5pt; }}
    .consent-summary ul {{ margin-bottom: 0; }}
    .signature-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 3mm; margin-top: 3mm; }}
    .signature-box {{ min-height: 20mm; padding: 2mm; border: 0.8pt solid #333333; }}
    .signature-line {{ margin-top: 5mm; border-bottom: 0.8pt solid #333333; height: 6mm; }}
    .print-note {{ margin-top: 2mm; font-size: 8.5pt; }}
    
    /* 改ページ制御の最適化: セクション全体のbreak-insideを解除し、子要素と見出し泣き別れを制御 */
    .section-block {{ break-inside: auto; page-break-inside: auto; padding: 0 0 4mm; border: 0; }}
    p, li, .note, .urgent {{ break-inside: avoid; page-break-inside: avoid; }}
    
    /* 紙節約用の2段組み設定 */
    .section-text {{ column-count: 2; column-gap: 6mm; }}
    h2, h3 {{ column-span: all; }}
    
    .doc-cover {{ padding-top: 0; background: #ffffff; }}
    .doc-cover .section-text {{ column-count: 1; }} /* 表紙セクションは2段組みにしない */
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
    
    .section-layout {{ display: grid; grid-template-columns: minmax(0, 1fr) 35mm; gap: 4mm; }}
    .section-visual {{ min-height: 25mm; padding: 1.5mm; border: 0.8pt solid #aaaaaa; background: #ffffff !important; }}
    .section-visual svg {{ max-width: 30mm; }}
    .section-visual img {{ max-width: 32mm; }}
    .section-visual figcaption {{ font-size: 7.5pt; }}
    
    .note, .urgent {{
      padding: 2mm;
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
  <style id="print-infographic-css" media="not all">
    @page {{ size: A4; margin: 8mm 8mm 8mm; }}
    * {{ box-shadow: none !important; text-shadow: none !important; }}
    html, body {{
      margin: 0; padding: 0;
      background: #ffffff !important; color: #111111;
      font-family: var(--font);
      font-size: 10pt; line-height: 1.25;
    }}
    .document, article {{ width: auto; margin: 0; padding: 0; border: 0; background: #ffffff; }}
    .print-consent {{ display: block; padding: 0 0 3mm; border-bottom: 1pt solid #111111; margin-bottom: 4mm; }}
    .print-consent h1 {{ margin: 0 0 2mm; font-size: 15pt; text-align: center; }}
    .consent-meta-grid {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 1.5mm; margin-top: 2mm; }}
    .consent-field {{ min-height: 8mm; padding: 1mm; border: 0.5pt solid #333333; font-size: 8pt; }}
    .consent-label {{ display: block; margin-bottom: 0.5mm; font-size: 7pt; font-weight: 700; }}
    .consent-summary {{ display: none; }} /* A4チラシでは説明本文と重複するため同意欄のサマリーは非表示 */
    .signature-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 3mm; margin-top: 3mm; }}
    .signature-box {{ min-height: 15mm; padding: 1.5mm; border: 0.8pt solid #333333; font-size: 8pt; }}
    .signature-line {{ margin-top: 3mm; border-bottom: 0.8pt solid #333333; height: 4mm; }}
    .print-note {{ margin-top: 1mm; font-size: 7.5pt; }}

    .section-block {{ break-inside: avoid; page-break-inside: avoid; padding: 0 0 3mm; border: 0; border-bottom: 0.5pt dashed #cccccc; }}
    .section-block:last-child {{ border-bottom: 0; }}
    .doc-cover {{ padding-top: 0; background: #ffffff; }}
    .eyebrow {{ margin: 0 0 1mm; font-size: 8.5pt; font-weight: 700; color: var(--accent-strong); }}
    h1, h2, h3 {{ break-after: avoid; page-break-after: avoid; color: #111111; line-height: 1.15; }}
    h1 {{ margin: 0 0 2mm; font-size: 16pt; }}
    h2 {{ margin: 0 0 1.5mm; padding: 0 0 0.5mm; border: 0; border-left: 3.5pt solid #111111; padding-left: 6px; font-size: 11pt; }}
    h3 {{ margin: 0 0 1mm; font-size: 9.5pt; }}
    p, ul, ol, dl {{ margin-top: 0; margin-bottom: 1.5mm; }}
    ul, ol {{ padding-left: 1.2em; }}
    li {{ margin: 0.5mm 0; }}
    a {{ color: #111111; text-decoration: none; }}
    .lead {{ font-size: 10pt; margin-top: 1mm; margin-bottom: 2mm; }}
    .note, .urgent {{
      break-inside: avoid; page-break-inside: avoid; padding: 1.5mm;
      border: 0.8pt solid #333333; background: #ffffff !important;
      font-size: 9pt; margin-bottom: 1.5mm;
    }}
    .step-list, .risk-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 1.5mm; margin-bottom: 1.5mm; }}
    .step-list section, .risk-grid section {{
      break-inside: avoid; page-break-inside: avoid; padding: 1.5mm;
      border: 0.5pt solid #666666; background: #ffffff !important;
      font-size: 9pt;
    }}
    .check-list {{ list-style: square; padding-left: 1.2em; }}
    .check-list li {{ padding-left: 0; }}
    .check-list li::before {{ display: none; }}

    .section-layout {{ display: grid; grid-template-columns: minmax(0, 1fr) 32mm; gap: 3mm; align-items: start; }}
    .section-visual {{ min-height: 22mm; padding: 1mm; border: 0.5pt solid #aaaaaa; background: #ffffff !important; }}
    .section-visual svg {{ max-width: 30mm; }}
    .section-visual img {{ max-width: 30mm; }}
    .section-visual figcaption {{ font-size: 7.5pt; margin-top: 3px; }}
  </style>
  <style id="slides-css" media="not all">
    body[data-mode="slides"] {{ overflow: hidden; background: #0b1329; }}
    body[data-mode="slides"] .app-toolbar {{
      background: rgba(11, 19, 41, 0.95); border-bottom-color: rgba(255, 255, 255, 0.12);
      backdrop-filter: blur(12px);
    }}
    body[data-mode="slides"] .toolbar-title {{ color: #81e6d9; font-weight: 600; }}
    body[data-mode="slides"] button {{
      border-color: rgba(255, 255, 255, 0.18); background: rgba(255, 255, 255, 0.06); color: #e2e8f0;
      border-radius: 8px; transition: all 0.25s ease;
    }}
    body[data-mode="slides"] button:hover {{
      background: rgba(255, 255, 255, 0.12); color: #ffffff; border-color: rgba(255, 255, 255, 0.3);
    }}
    body[data-mode="slides"] button[aria-pressed="true"] {{
      background: #0d9488; color: #ffffff; border-color: #0d9488; font-weight: 600;
      box-shadow: 0 0 12px rgba(13, 148, 136, 0.4);
    }}
    body[data-mode="slides"] .document {{
      width: 100vw; height: calc(100vh - 59px); margin: 0; padding: 0;
      overflow-x: auto; overflow-y: hidden; scroll-snap-type: x mandatory;
    }}
    body[data-mode="slides"] article {{
      display: flex; width: max-content; height: 100%; border: 0; box-shadow: none; background: transparent;
    }}
    body[data-mode="slides"] .section-block {{
      display: flex; flex-direction: column; justify-content: center;
      width: 100vw; height: 100%; padding: clamp(40px, 8vw, 100px);
      overflow-y: auto; scroll-snap-align: start; border: 0;
      background: linear-gradient(135deg, #f0fdf4 0%, #e0f2fe 100%);
      box-shadow: inset 0 0 80px rgba(15, 118, 110, 0.03);
    }}
    body[data-mode="slides"] .section-layout {{
      grid-template-columns: minmax(0, 1fr) minmax(320px, 36vw);
      gap: clamp(32px, 6vw, 88px);
      align-items: center;
      min-height: 65vh;
    }}
    body[data-mode="slides"] .section-visual {{
      min-height: min(46vh, 380px);
      border: 1px solid rgba(20, 184, 166, 0.12);
      border-radius: 20px;
      background: #ffffff;
      padding: clamp(20px, 3vw, 36px);
      box-shadow: 0 20px 40px -12px rgba(15, 118, 110, 0.08), 0 1px 3px rgba(15, 118, 110, 0.02);
    }}
    body[data-mode="slides"] .section-visual svg {{ max-width: 340px; }}
    body[data-mode="slides"] .section-visual img {{
      max-width: min(540px, 100%);
      border-radius: 12px;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.04);
    }}
    body[data-mode="slides"] .section-visual figcaption {{
      font-size: clamp(13px, 1.1vw, 16px);
      color: #475569;
      background: #f1f5f9;
      padding: 6px 12px;
      border-radius: 6px;
      display: inline-block;
      margin-top: 14px;
    }}
    body[data-mode="slides"] .doc-cover {{
      background: linear-gradient(135deg, #e6fffa 0%, #e0f2fe 100%);
    }}
    body[data-mode="slides"] .print-consent {{ display: none !important; }}
    body[data-mode="slides"] h1 {{
      max-width: 1100px;
      font-size: clamp(38px, 5.2vw, 68px);
      font-weight: 800;
      color: #0f172a;
      line-height: 1.25;
      letter-spacing: -0.02em;
    }}
    body[data-mode="slides"] h2 {{
      margin-bottom: clamp(24px, 3.5vw, 42px);
      padding-left: 20px;
      border-left: 8px solid #0d9488;
      border-radius: 4px;
      font-size: clamp(30px, 3.8vw, 46px);
      font-weight: 700;
      color: #0f172a;
      line-height: 1.3;
      letter-spacing: -0.01em;
    }}
    body[data-mode="slides"] .section-block.has-slide-summary .section-text > :not(h2):not(.slide-summary) {{ display: none; }}
    body[data-mode="slides"] .slide-summary {{ display: block; font-weight: 500; color: #0f766e; }}
    body[data-mode="slides"] h3 {{
      font-size: clamp(20px, 2.2vw, 28px);
      color: #0d9488;
      margin-top: 24px;
      margin-bottom: 12px;
      font-weight: 700;
    }}
    body[data-mode="slides"] p, body[data-mode="slides"] li {{
      max-width: 1100px;
      font-size: clamp(19px, 1.8vw, 26px);
      line-height: 1.6;
      color: #334155;
    }}
    body[data-mode="slides"] li {{
      position: relative;
      padding-left: 1.6em;
      list-style-type: none;
      margin: 10px 0;
    }}
    body[data-mode="slides"] li::before {{
      content: "•";
      color: #0d9488;
      font-weight: bold;
      display: inline-block;
      width: 1em;
      margin-left: -1em;
      font-size: 1.3em;
      vertical-align: middle;
      position: absolute;
      left: 0.8em;
      top: -0.05em;
    }}
    body[data-mode="slides"] .lead {{
      font-size: clamp(22px, 2.2vw, 30px);
      line-height: 1.65;
      color: #334155;
      font-weight: 500;
    }}
    body[data-mode="slides"] .step-list, body[data-mode="slides"] .risk-grid {{
      grid-template-columns: repeat(2, minmax(320px, 1fr)); gap: 20px;
      margin-top: 18px;
    }}
    body[data-mode="slides"] .note, body[data-mode="slides"] .urgent,
    body[data-mode="slides"] .step-list section, body[data-mode="slides"] .risk-grid section {{
      background: rgba(255, 255, 255, 0.85);
      border: 1px solid rgba(20, 184, 166, 0.12);
      border-radius: 16px;
      padding: clamp(20px, 2.5vw, 32px);
      box-shadow: 0 10px 30px -10px rgba(15, 118, 110, 0.06), 0 1px 3px rgba(15, 118, 110, 0.02);
      backdrop-filter: blur(12px);
      transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease;
    }}
    body[data-mode="slides"] .step-list section:hover,
    body[data-mode="slides"] .risk-grid section:hover {{
      transform: translateY(-4px);
      box-shadow: 0 20px 40px -15px rgba(15, 118, 110, 0.12);
      border-color: rgba(20, 184, 166, 0.3);
    }}
    body[data-mode="slides"] .note {{
      border-left: 8px solid #d97706;
      background: linear-gradient(135deg, rgba(255, 255, 255, 0.85) 0%, rgba(254, 243, 199, 0.4) 100%);
    }}
    body[data-mode="slides"] .urgent {{
      border-left: 8px solid #b42318;
      background: linear-gradient(135deg, rgba(255, 255, 255, 0.85) 0%, rgba(254, 226, 226, 0.4) 100%);
      border-color: rgba(180, 35, 24, 0.12);
    }}
    .slide-pagination {{ display: none; }}
    body[data-mode="slides"] .slide-pagination {{
      position: fixed; right: 30px; bottom: 24px; z-index: 20;
      display: flex; align-items: center; gap: 16px; min-width: 168px; justify-content: center;
      padding: 10px 16px; border: 1px solid rgba(255, 255, 255, 0.15); border-radius: 30px;
      background: rgba(11, 19, 41, 0.88); color: #f8fafc; font-size: 15px; line-height: 1;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.25); backdrop-filter: blur(10px);
    }}
    body[data-mode="slides"] .slide-nav {{
      display: inline-flex; align-items: center; justify-content: center;
      width: 36px; height: 36px; min-height: 36px; padding: 0;
      border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 50%;
      background: rgba(255, 255, 255, 0.1); color: #ffffff; font-size: 20px; line-height: 1;
      cursor: pointer; transition: all 0.2s ease;
    }}
    body[data-mode="slides"] .slide-nav:hover {{
      background: #0d9488; color: #ffffff; border-color: #0d9488;
      box-shadow: 0 0 8px rgba(13, 148, 136, 0.5);
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
      <button type="button" data-mode="infographic" aria-pressed="false">A4チラシ</button>
      <button type="button" data-mode="slides" aria-pressed="false">スライド</button>
      <button type="button" data-action="print">印刷/PDF</button>
    </div>
  </header>

  <main class="document" id="top">
    <article>
{print_consent}
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
    const printInfographic = document.getElementById("print-infographic-css");
    const slides = document.getElementById("slides-css");
    const buttons = document.querySelectorAll("[data-mode]");
    const documentFrame = document.querySelector(".document");
    const slideBlocks = Array.from(document.querySelectorAll(".section-block"));
    const currentSlide = document.getElementById("current-slide");
    const totalSlides = document.getElementById("total-slides");
    let activeSlideIndex = 0;

    function setMode(mode) {{
      document.body.dataset.mode = mode;
      printLarge.media = mode === "large" ? "all" : (mode === "screen" || mode === "slides" ? "print" : "not all");
      printCompact.media = mode === "compact" ? "all" : "not all";
      printInfographic.media = mode === "infographic" ? "all" : "not all";
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
        "version": "",
        "diagnosis": "",
        "procedure_name": "",
        "anesthesia": "",
        "laterality": "",
        "document_type": "",
        "print_consent": "true",
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
    metadata["images"] = ", ".join(image_paths)
    metadata["build_datetime"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    body_lines = lines[start_idx:]

    # 2. 最初にあるH1 (# タイトル) を抽出し、フロントマターにタイトルがない場合はそれをタイトルにする
    content_lines = []
    for line in body_lines:
        if line.startswith("# ") or line.startswith("# **"):
            # H1は表紙タイトルとして扱い、本文には重複表示しない
            if title == "患者説明資料":
                raw_title = line.replace("#", "").replace("**", "").strip()
                if raw_title:
                    title = raw_title
            continue
        content_lines.append(line)

    # 3. 行ごとに解析してHTMLセクションに分類
    sections = []
    current_section_content = []
    current_section_title = ""
    current_section_visual_key = ""
    current_section_slide_summary = ""
    is_cover = True  # 最初のH2に到達するまではカバーセクション

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
                sections.append(
                    {
                        "is_cover": True,
                        "title": title,
                        "visual_key": "cover",
                        "slide_summary": current_section_slide_summary,
                        "content": "\n".join(current_section_content),
                    }
                )
                is_cover = False
            else:
                # 通常セクションを保存
                sections.append(
                    {
                        "is_cover": False,
                        "title": current_section_title,
                        "visual_key": current_section_visual_key,
                        "slide_summary": current_section_slide_summary,
                        "content": "\n".join(current_section_content),
                    }
                )

            # 新しいセクションをスタート
            current_section_content = []
            current_section_title = h2_text
            current_section_visual_key = ""
            current_section_slide_summary = ""
            current_section_content.append(f"      <h2>{h2_text}</h2>")
            continue

        visual_match = re.match(
            r"^\{\{visual:\s*([a-zA-Z0-9_-]+)\s*\}\}$", cleaned_line
        )
        if visual_match:
            current_section_visual_key = visual_match.group(1)
            continue

        slide_summary_match = re.match(
            r"^\{\{slide_summary:\s*(.*?)\s*\}\}$", cleaned_line
        )
        if slide_summary_match:
            current_section_slide_summary = format_inline_elements(
                slide_summary_match.group(1).strip()
            )
            continue

        # 空行
        if not cleaned_line:
            if in_p:
                current_section_content.append("        </p>")
                in_p = False
            continue

        # リストマークアップのクローズ判定
        is_list_line = (
            cleaned_line.startswith("- ")
            or cleaned_line.startswith("* ")
            or cleaned_line.startswith("+ ")
            or re.match(r"^\d+\.\s+", cleaned_line)
        )

        if not is_list_line:
            if in_ul:
                current_section_content.append("        </ul>")
                in_ul = False
            if in_ol:
                current_section_content.append("        </ol>")
                in_ol = False

        # 箇条書きリスト (UL)
        if (
            cleaned_line.startswith("- ")
            or cleaned_line.startswith("* ")
            or cleaned_line.startswith("+ ")
        ):
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
            is_warning = (
                quote_text.startswith("[!NOTE]")
                or quote_text.startswith("[!TIP]")
                or quote_text.startswith("[!IMPORTANT]")
            )
            is_urgent = quote_text.startswith("[!WARNING]") or quote_text.startswith(
                "[!CAUTION]"
            )

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
        sections.append(
            {
                "is_cover": True,
                "title": title,
                "visual_key": "cover",
                "slide_summary": current_section_slide_summary,
                "content": "\n".join(current_section_content),
            }
        )
    else:
        sections.append(
            {
                "is_cover": False,
                "title": current_section_title,
                "visual_key": current_section_visual_key,
                "slide_summary": current_section_slide_summary,
                "content": "\n".join(current_section_content),
            }
        )

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


def section_visual_svg(visual_key):
    """Markdownで指定された図版キーに対応する説明用模式図を返す。"""
    visuals = {
        "cover": """<svg viewBox="0 0 360 260" role="img" aria-label="耳の構造と真珠腫説明資料の表紙図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <path d="M42 148c28-68 84-104 158-94 62 8 101 49 110 113" fill="none" stroke="#bfdbfe" stroke-width="22" stroke-linecap="round"/>
  <path d="M92 156c44-28 93-32 144-10" fill="none" stroke="#94a3b8" stroke-width="13" stroke-linecap="round"/>
  <circle cx="223" cy="116" r="32" fill="#fde68a" stroke="#d97706" stroke-width="6"/>
  <text x="223" y="121" text-anchor="middle" font-size="15" fill="#78350f">真珠腫</text>
  <text x="180" y="226" text-anchor="middle" font-size="18" fill="#0f766e">耳の奥を見ながら説明します</text>
</svg>""",
        "cholesteatoma-growth": """<svg viewBox="0 0 360 260" role="img" aria-label="鼓膜のへこみから真珠腫が耳小骨へ広がる図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <text x="38" y="34" font-size="16" fill="#0f766e">外耳道</text>
  <text x="167" y="34" font-size="16" fill="#0f766e">鼓膜</text>
  <text x="258" y="34" font-size="16" fill="#0f766e">中耳</text>
  <path d="M34 132h104" stroke="#93c5fd" stroke-width="34" stroke-linecap="round"/>
  <path d="M152 72c28 36 28 82 0 118" fill="none" stroke="#2563eb" stroke-width="10" stroke-linecap="round"/>
  <path d="M159 96c30 18 50 28 77 28" fill="none" stroke="#64748b" stroke-width="8" stroke-linecap="round"/>
  <circle cx="235" cy="124" r="31" fill="#fde68a" stroke="#d97706" stroke-width="6"/>
  <path d="M247 116c31-11 51-8 68 9" fill="none" stroke="#dc2626" stroke-width="5" stroke-linecap="round"/>
  <text x="236" y="130" text-anchor="middle" font-size="14" fill="#78350f">真珠腫</text>
  <text x="278" y="160" text-anchor="middle" font-size="15" fill="#dc2626">骨を溶かす</text>
</svg>""",
        "surgery-purpose": """<svg viewBox="0 0 360 260" role="img" aria-label="真珠腫を取り除き合併症を防ぐ図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <circle cx="112" cy="122" r="45" fill="#fde68a" stroke="#d97706" stroke-width="6"/>
  <text x="112" y="127" text-anchor="middle" font-size="16" fill="#78350f">真珠腫</text>
  <path d="M168 122h58" stroke="#0f766e" stroke-width="9" stroke-linecap="round"/>
  <path d="M211 99l26 23-26 23" fill="none" stroke="#0f766e" stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/>
  <rect x="248" y="77" width="68" height="90" rx="12" fill="#dcfce7" stroke="#16a34a" stroke-width="6"/>
  <path d="M266 124l14 14 26-36" fill="none" stroke="#16a34a" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>
  <text x="180" y="215" text-anchor="middle" font-size="18" fill="#0f766e">取り除いて進行を止める</text>
</svg>""",
        "surgical-approaches": """<svg viewBox="0 0 360 260" role="img" aria-label="真珠腫の広がりに応じた術式選択図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <g font-size="13" fill="#0f172a" text-anchor="middle">
    <rect x="22" y="56" width="72" height="104" rx="10" fill="#eff6ff" stroke="#2563eb" stroke-width="4"/>
    <text x="58" y="48">浅い</text><path d="M40 118h36" stroke="#2563eb" stroke-width="9" stroke-linecap="round"/>
    <rect x="106" y="56" width="72" height="104" rx="10" fill="#ecfeff" stroke="#0891b2" stroke-width="4"/>
    <text x="142" y="48">奥へ</text><path d="M124 118h36" stroke="#0891b2" stroke-width="9" stroke-linecap="round"/><circle cx="158" cy="112" r="13" fill="#fde68a" stroke="#d97706" stroke-width="4"/>
    <rect x="190" y="56" width="72" height="104" rx="10" fill="#f0fdf4" stroke="#16a34a" stroke-width="4"/>
    <text x="226" y="48">壁を再建</text><path d="M204 118h44" stroke="#16a34a" stroke-width="9" stroke-linecap="round"/><path d="M226 82v70" stroke="#64748b" stroke-width="5"/>
    <rect x="274" y="56" width="72" height="104" rx="10" fill="#fff7ed" stroke="#ea580c" stroke-width="4"/>
    <text x="310" y="48">開放型</text><path d="M288 118h44" stroke="#ea580c" stroke-width="9" stroke-linecap="round"/><path d="M314 84c17 26 17 53 0 79" fill="none" stroke="#ea580c" stroke-width="5"/>
  </g>
  <text x="180" y="215" text-anchor="middle" font-size="17" fill="#0f766e">広がりで安全な方法を選ぶ</text>
</svg>""",
        "reconstruction": """<svg viewBox="0 0 360 260" role="img" aria-label="鼓膜と耳小骨を軟骨や筋膜で再建する図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <path d="M70 132h92" stroke="#93c5fd" stroke-width="30" stroke-linecap="round"/>
  <path d="M174 76c30 35 30 78 0 113" fill="none" stroke="#2563eb" stroke-width="9" stroke-linecap="round"/>
  <circle cx="210" cy="125" r="14" fill="#cbd5e1" stroke="#64748b" stroke-width="5"/>
  <circle cx="246" cy="125" r="14" fill="#cbd5e1" stroke="#64748b" stroke-width="5"/>
  <path d="M211 125h35" stroke="#64748b" stroke-width="6"/>
  <path d="M159 111l47-22 14 26-48 22z" fill="#bbf7d0" stroke="#16a34a" stroke-width="5"/>
  <text x="192" y="79" text-anchor="middle" font-size="14" fill="#166534">軟骨・筋膜</text>
  <text x="232" y="160" text-anchor="middle" font-size="14" fill="#475569">耳小骨</text>
  <text x="180" y="215" text-anchor="middle" font-size="17" fill="#0f766e">必要なら聞こえの道を作り直す</text>
</svg>""",
        "recurrence-types": """<svg viewBox="0 0 360 260" role="img" aria-label="遺残と再形成の2種類の再発を示す図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <rect x="28" y="54" width="138" height="124" rx="14" fill="#fefce8" stroke="#ca8a04" stroke-width="5"/>
  <text x="97" y="82" text-anchor="middle" font-size="18" fill="#854d0e">遺残</text>
  <circle cx="97" cy="121" r="10" fill="#fde68a" stroke="#d97706" stroke-width="5"/>
  <path d="M97 137c0 18 0 28 0 42" stroke="#d97706" stroke-width="5" stroke-dasharray="6 6"/>
  <rect x="194" y="54" width="138" height="124" rx="14" fill="#eff6ff" stroke="#2563eb" stroke-width="5"/>
  <text x="263" y="82" text-anchor="middle" font-size="18" fill="#1d4ed8">再形成</text>
  <path d="M232 109c25 4 46 2 65-8" fill="none" stroke="#2563eb" stroke-width="8" stroke-linecap="round"/>
  <path d="M267 104c-18 18-23 37-13 57" fill="none" stroke="#d97706" stroke-width="7" stroke-linecap="round"/>
  <text x="180" y="220" text-anchor="middle" font-size="17" fill="#0f766e">種類が違うため長期通院が必要</text>
</svg>""",
        "sniffing-pressure": """<svg viewBox="0 0 360 260" role="img" aria-label="鼻すすりで鼓膜が内側へ引き込まれる図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <text x="84" y="56" text-anchor="middle" font-size="15" fill="#0f766e">鼻すすり</text>
  <path d="M95 78c45 0 72 19 94 54" fill="none" stroke="#0284c7" stroke-width="8" stroke-linecap="round" stroke-dasharray="8 8"/>
  <path d="M181 111l12 26-29-2" fill="#0284c7"/>
  <path d="M215 72c31 40 31 82 0 122" fill="none" stroke="#2563eb" stroke-width="10" stroke-linecap="round"/>
  <path d="M214 96c-32 17-45 40-40 69" fill="none" stroke="#dc2626" stroke-width="8" stroke-linecap="round"/>
  <text x="255" y="132" font-size="15" fill="#dc2626">鼓膜がへこむ</text>
  <text x="180" y="220" text-anchor="middle" font-size="17" fill="#0f766e">陰圧が再形成の原因になる</text>
</svg>""",
        "complications": """<svg viewBox="0 0 360 260" role="img" aria-label="耳の周囲にある顔面神経や内耳へのリスク図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <path d="M78 137c31-58 76-82 134-66 42 12 65 44 69 91" fill="none" stroke="#bfdbfe" stroke-width="20" stroke-linecap="round"/>
  <circle cx="210" cy="125" r="25" fill="#fde68a" stroke="#d97706" stroke-width="5"/>
  <path d="M254 84c29 26 30 62 2 91" fill="none" stroke="#dc2626" stroke-width="8" stroke-linecap="round"/>
  <text x="279" y="80" font-size="14" fill="#dc2626">顔面神経</text>
  <path d="M286 138c20-17 34-16 43 2s-2 34-24 34" fill="none" stroke="#7c3aed" stroke-width="7" stroke-linecap="round"/>
  <text x="282" y="199" font-size="14" fill="#6d28d9">内耳</text>
  <text x="180" y="226" text-anchor="middle" font-size="17" fill="#0f766e">近くの神経・内耳に注意</text>
</svg>""",
        "red-flags": """<svg viewBox="0 0 360 260" role="img" aria-label="退院後すぐ連絡すべき赤信号症状の図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <path d="M180 42l88 154H92z" fill="#fee2e2" stroke="#dc2626" stroke-width="8" stroke-linejoin="round"/>
  <path d="M180 95v54" stroke="#dc2626" stroke-width="13" stroke-linecap="round"/>
  <circle cx="180" cy="171" r="8" fill="#dc2626"/>
  <g font-size="14" fill="#991b1b">
    <text x="35" y="68">強いめまい</text>
    <text x="253" y="68">顔の麻痺</text>
    <text x="36" y="218">高熱・頭痛</text>
    <text x="245" y="218">透明な耳だれ</text>
  </g>
</svg>""",
        "follow-up": """<svg viewBox="0 0 360 260" role="img" aria-label="術後の耳の安静と定期通院の図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <rect x="68" y="58" width="224" height="142" rx="14" fill="#f8fafc" stroke="#94a3b8" stroke-width="5"/>
  <path d="M96 101h168M96 141h168" stroke="#cbd5e1" stroke-width="5"/>
  <circle cx="118" cy="101" r="13" fill="#ccfbf1" stroke="#0f766e" stroke-width="5"/>
  <path d="M111 101l6 7 13-18" fill="none" stroke="#0f766e" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="118" cy="141" r="13" fill="#ccfbf1" stroke="#0f766e" stroke-width="5"/>
  <path d="M111 141l6 7 13-18" fill="none" stroke="#0f766e" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>
  <text x="180" y="226" text-anchor="middle" font-size="17" fill="#0f766e">安静と定期チェック</text>
</svg>""",
        "child-ear-growth": """<svg viewBox="0 0 360 260" role="img" aria-label="小児の乳突蜂巣の発育と中耳炎の影響を示す図">
  <rect width="360" height="260" rx="18" fill="#ffffff"/>
  <text x="100" y="50" text-anchor="middle" font-size="15" fill="#0f766e">発育良好</text>
  <circle cx="100" cy="119" r="55" fill="#dbeafe" stroke="#2563eb" stroke-width="5"/>
  <g fill="#ffffff" stroke="#60a5fa" stroke-width="3">
    <circle cx="79" cy="105" r="10"/><circle cx="105" cy="93" r="12"/><circle cx="123" cy="120" r="11"/><circle cx="91" cy="139" r="12"/>
  </g>
  <text x="260" y="50" text-anchor="middle" font-size="15" fill="#b45309">発育不良</text>
  <circle cx="260" cy="119" r="55" fill="#ffedd5" stroke="#ea580c" stroke-width="5"/>
  <path d="M226 119h68" stroke="#ea580c" stroke-width="13" stroke-linecap="round"/>
  <text x="180" y="226" text-anchor="middle" font-size="17" fill="#0f766e">子どもは耳の成長も守る</text>
</svg>""",
    }
    return visuals.get(visual_key, "")


def section_visual_html(visual_key):
    """図版キーに対応するドラフト画像を返す。画像は医師レビュー前のサンプル扱い。"""
    visuals = {
        "cover": ("cholesteatoma-growth.png", "真珠腫が中耳で広がる位置関係の模式図"),
        "cholesteatoma-growth": (
            "cholesteatoma-growth.png",
            "鼓膜の奥に真珠腫ができ、耳小骨や周囲の骨へ近づく様子",
        ),
        "surgery-purpose": (
            "surgery-purpose.png",
            "耳の後ろ側から病変へ到達し、真珠腫を取り除く考え方",
        ),
        "surgical-approaches": (
            "surgery-purpose.png",
            "真珠腫の広がりに合わせて安全な手術方法を選びます",
        ),
        "reconstruction": (
            "reconstruction.png",
            "鼓膜や耳小骨を、軟骨や筋膜で補う再建イメージ",
        ),
        "recurrence-types": (
            "recurrence-types.png",
            "遺残と再形成という2種類の再発パターン",
        ),
        "sniffing-pressure": (
            "sniffing-pressure.png",
            "鼻すすりで耳の中に陰圧がかかり、鼓膜が奥へ引かれる仕組み",
        ),
        "complications": (
            "complications-risk.png",
            "顔面神経や内耳など、真珠腫の近くにある重要な構造",
        ),
        "red-flags": (
            "complications-risk.png",
            "強いめまい、顔の動きにくさ、聞こえの急な悪化などに注意します",
        ),
        "follow-up": (
            "recurrence-types.png",
            "手術後も再発確認のため、長期間の定期通院が必要です",
        ),
        "child-ear-growth": (
            "cholesteatoma-growth.png",
            "小児では耳の成長も考えて治療方針を決めます",
        ),
    }
    item = visuals.get(visual_key)
    if not item:
        return ""
    filename, caption = item
    src = f"../images/draft/cholesteatoma/{filename}"
    return f"""<figure>
  <img src="{html.escape(src)}" alt="{html.escape(caption)}">
  <figcaption>{html.escape(caption)}<br>※AI生成ドラフト。臨床使用前に医師の確認が必要です。</figcaption>
</figure>"""


def generate_print_consent_html(title, metadata):
    """印刷時だけ表示する手術説明書・同意確認欄を生成する。"""
    document_type = metadata.get("document_type") or "手術説明書・同意確認書"
    diagnosis = metadata.get("diagnosis") or ""
    procedure_name = metadata.get("procedure_name") or title
    anesthesia = metadata.get("anesthesia") or ""
    laterality = metadata.get("laterality") or "右・左"
    version = metadata.get("version") or ""
    updated = metadata.get("updated") or ""
    reviewed_by = metadata.get("reviewed_by") or "未確認"
    review_date = metadata.get("review_date") or ""

    return f"""      <section class="print-consent" aria-label="印刷用同意確認欄">
        <h1>{html.escape(document_type)}</h1>
        <div class="consent-meta-grid">
          <div class="consent-field"><span class="consent-label">文書名</span>{html.escape(title)}</div>
          <div class="consent-field"><span class="consent-label">病名</span>{html.escape(diagnosis)}</div>
          <div class="consent-field"><span class="consent-label">手術名</span>{html.escape(procedure_name)}</div>
          <div class="consent-field"><span class="consent-label">麻酔</span>{html.escape(anesthesia)}</div>
          <div class="consent-field"><span class="consent-label">術側</span>{html.escape(laterality)}</div>
          <div class="consent-field"><span class="consent-label">手術予定日</span>　　　　年　　　月　　　日</div>
          <div class="consent-field"><span class="consent-label">患者氏名</span></div>
          <div class="consent-field"><span class="consent-label">説明日</span>　　　　年　　　月　　　日</div>
          <div class="consent-field"><span class="consent-label">説明医</span></div>
          <div class="consent-field"><span class="consent-label">同席者・ご家族</span></div>
        </div>
        <div class="consent-summary">
          <h2>説明・確認項目</h2>
          <ul>
            <li>病気の状態と、手術が必要となる理由</li>
            <li>手術の目的、方法、麻酔、術式が変更となる可能性</li>
            <li>期待される効果と、聴力が改善しないまたは悪化する可能性</li>
            <li>出血、感染、めまい、難聴、耳鳴り、味覚障害、顔面神経麻痺、髄液漏、髄膜炎、再発などの合併症</li>
            <li>経過観察や処置など、手術以外の選択肢と限界</li>
            <li>術後の注意点、受診が必要な症状、長期通院の必要性</li>
          </ul>
        </div>
        <div class="signature-grid">
          <div class="signature-box">
            <strong>患者さん・代諾者署名</strong>
            <div class="signature-line"></div>
            <div>続柄：</div>
          </div>
          <div class="signature-box">
            <strong>医療者記入欄</strong>
            <div class="signature-line"></div>
            <div>説明医署名：</div>
          </div>
        </div>
        <p class="print-note">この欄は印刷時の確認用です。電子表示・スライド表示では説明を見やすくするため、本文中心の構成にしています。版数: {html.escape(version)} / 更新日: {html.escape(updated)} / 医師レビュー: {html.escape(reviewed_by)} {html.escape(review_date)}</p>
      </section>"""


# セクションのHTMLブロックを生成
def generate_sections_html(title, description, sections):
    html_blocks = []

    for sec in sections:
        slide_summary = sec.get("slide_summary", "")
        slide_summary_html = (
            f'            <p class="slide-summary">{slide_summary}</p>'
            if slide_summary
            else ""
        )
        summary_class = " has-slide-summary" if slide_summary else ""
        if sec["is_cover"]:
            # カバーセクションのマークアップ
            lead_html = (
                f'        <p class="lead">{description}</p>' if description else ""
            )
            visual_html = section_visual_html(sec.get("visual_key", "cover"))
            block = f"""      <section class="doc-cover section-block{summary_class}">
        <div class="section-layout">
          <div class="section-text">
            <p class="eyebrow">患者さん・ご家族への説明資料</p>
            <h1>{title}</h1>
{lead_html}
{sec["content"]}
{slide_summary_html}
          </div>
          <div class="section-visual" aria-hidden="true">
{visual_html}
          </div>
        </div>
      </section>"""
        else:
            # 通常セクションのマークアップ
            visual_html = section_visual_html(sec.get("visual_key", ""))
            block = f"""      <section class="section-block{summary_class}">
        <div class="section-layout">
          <div class="section-text">
{sec["content"]}
{slide_summary_html}
          </div>
          <div class="section-visual" aria-hidden="true">
{visual_html}
          </div>
        </div>
      </section>"""
        html_blocks.append(block)

    return "\n\n".join(html_blocks)


# メイン処理：明示されたMarkdownファイルのみ変換
def convert_all_markdowns(target_names):
    if not target_names:
        print("ERROR: 変換するMarkdownファイルを1つ以上指定してください。")
        return 1

    excluded_filenames = {
        "README.md",
        "CHANGELOG.md",
        "rules.md",
        "Gitプロジェクト一覧.md",
        "Patient-information.md",
    }
    normalized_targets = []
    seen_targets = set()
    for target_name in target_names:
        filename = os.path.basename(target_name)
        if not filename.endswith(".md"):
            filename = f"{filename}.md"
        if filename not in seen_targets:
            normalized_targets.append(filename)
            seen_targets.add(filename)

    excluded_targets = [
        filename for filename in normalized_targets if filename in excluded_filenames
    ]
    if excluded_targets:
        print("ERROR: 以下のファイルはHTML変換対象外です。")
        for filename in excluded_targets:
            print(f" - {filename}")
        return 1

    md_files = []
    missing_targets = []
    for filename in normalized_targets:
        filepath = os.path.join(base_dir, "src", filename)
        if os.path.exists(filepath):
            md_files.append(filepath)
        else:
            missing_targets.append(filename)

    if missing_targets:
        print("ERROR: 以下のMarkdownファイルがsrc配下に見つかりません。")
        for filename in missing_targets:
            print(f" - {filename}")
        return 1

    success_count = 0
    error_files = []

    print(f"Total target markdown files: {len(md_files)}")

    for filepath in md_files:
        filename = os.path.basename(filepath)

        print(f"Converting {filename}...")
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                md_text = f.read()

            title, description, sections, metadata = parse_markdown_to_sections(md_text)
            content_html = generate_sections_html(title, description, sections)

            # print_consent が "false" の場合は、署名欄を出力せず空にする
            if metadata.get("print_consent", "true").lower() == "false":
                print_consent_html = ""
            else:
                print_consent_html = generate_print_consent_html(title, metadata)

            # メタデータコメント文字列の生成
            meta_comments_lines = [
                "<!--",
                f"  Meta:reviewed_by: {metadata.get('reviewed_by', '')}",
                f"  Meta:review_date: {metadata.get('review_date', '')}",
                f"  Meta:evidence_source: {metadata.get('evidence_source', '')}",
                f"  Meta:change_reason: {metadata.get('change_reason', '')}",
                f"  Meta:version: {metadata.get('version', '')}",
                f"  Meta:build_datetime: {metadata.get('build_datetime', '')}",
                f"  Meta:images: {metadata.get('images', '')}",
                "-->",
            ]
            meta_comments_str = "\n".join(meta_comments_lines)

            # HTMLベーステンプレートに流し込む
            final_html = html_template.format(
                title=title,
                content=content_html,
                print_consent=print_consent_html,
                meta_comments=meta_comments_str,
            )

            # 出力ファイル名を作成 (.md から .html)
            out_filename = os.path.splitext(filename)[0] + ".html"
            out_filepath = os.path.join(html_dir, out_filename)

            with open(out_filepath, "w", encoding="utf-8") as out_f:
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
        return 1

    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="患者説明文書MarkdownをHTMLに変換します。"
    )
    parser.add_argument(
        "targets",
        nargs="+",
        help="変換するMarkdownファイル名。必ず1つ以上指定してください。",
    )
    args = parser.parse_args()
    raise SystemExit(convert_all_markdowns(args.targets))
