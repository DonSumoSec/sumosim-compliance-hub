#!/usr/bin/env python3
"""
SumoSim Compliance Hub Generator
Reads SumoSim_Compliance_Hub.xlsx and generates a self-contained index.html
Run: python3 generate_hub.py
Output: index.html  (upload this to GitHub repo)
"""

import openpyxl
import json
from datetime import datetime

XL_FILE = "SumoSim_Compliance_Hub.xlsx"
OUTPUT_FILE = "index.html"

def load_data(xl_path):
    wb = openpyxl.load_workbook(xl_path, data_only=True)

    # --- Dashboard: documents ---
    ws = wb['📊 Dashboard']
    docs = []
    for row in ws.iter_rows(min_row=7, max_row=30, values_only=True):
        if row[1] and row[1] != 'DOCUMENT':
            docs.append({
                "title":       str(row[1]).strip(),
                "category":    str(row[2]).strip() if row[2] else "",
                "version":     str(row[3]).strip() if row[3] else "",
                "status":      str(row[4]).strip() if row[4] else "",
                "reviewedBy":  str(row[5]).strip() if row[5] else "",
                "lastReviewed":str(row[6]).strip() if row[6] else "",
                "nextReview":  str(row[7]).strip() if row[7] else "",
                "websitePage": str(row[8]).strip() if row[8] else "",
                "published":   str(row[9]).strip() if row[9] else "",
            })

    # --- Content Library ---
    ws2 = wb['📝 Content Library']
    content = {}
    for row in ws2.iter_rows(min_row=5, max_row=60, values_only=True):
        if row[1] and row[1] != 'DOCUMENT NAME':
            title = str(row[1]).strip()
            section = str(row[2]).strip() if row[2] else ""
            live = str(row[3]).strip() if row[3] and str(row[3]) != '—' else ""
            draft = str(row[4]).strip() if row[4] and str(row[4]) != '—' else ""
            notes = str(row[5]).strip() if row[5] else ""
            if title not in content:
                content[title] = []
            content[title].append({
                "section": section,
                "live": live,
                "draft": draft,
                "notes": notes
            })

    return docs, content


def generate_html(docs, content):
    generated = datetime.now().strftime("%d %b %Y %H:%M")

    # Stats
    live_count = sum(1 for d in docs if d['status'] == 'Live')
    draft_count = sum(1 for d in docs if d['status'] == 'Draft')
    review_count = sum(1 for d in docs if d['status'] == 'Under Review')
    published_count = sum(1 for d in docs if d['published'] == 'Yes')

    docs_json = json.dumps(docs)
    content_json = json.dumps(content)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SumoSim | Compliance Hub</title>
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background:#0a0a0a; color:#e0e0e0; min-height:100vh; }}
  .topbar {{ background:#111; border-bottom:1px solid #222; padding:12px 24px; display:flex; align-items:center; justify-content:space-between; }}
  .topbar-brand {{ font-size:15px; font-weight:600; letter-spacing:0.5px; }}
  .topbar-brand span {{ color:#fff; }}
  .topbar-brand .dot {{ color:#666; margin:0 8px; }}
  .topbar-meta {{ font-size:12px; color:#555; }}
  .layout {{ display:flex; height:calc(100vh - 45px); }}
  .sidebar {{ width:220px; background:#111; border-right:1px solid #1e1e1e; overflow-y:auto; flex-shrink:0; padding:16px 0; }}
  .sidebar-section {{ padding:8px 16px 4px; font-size:10px; font-weight:700; color:#444; letter-spacing:1.5px; text-transform:uppercase; }}
  .sidebar-item {{ padding:8px 16px; font-size:13px; color:#888; cursor:pointer; border-left:2px solid transparent; transition:all 0.15s; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  .sidebar-item:hover {{ color:#ccc; background:#161616; }}
  .sidebar-item.active {{ color:#fff; border-left-color:#fff; background:#161616; }}
  .sidebar-badge {{ display:inline-block; font-size:9px; padding:1px 5px; border-radius:3px; margin-left:6px; font-weight:600; vertical-align:middle; }}
  .badge-live {{ background:#1a3a1a; color:#4caf50; }}
  .badge-draft {{ background:#3a2a1a; color:#ff9800; }}
  .badge-review {{ background:#1a2a3a; color:#2196f3; }}
  .main {{ flex:1; overflow-y:auto; padding:24px; }}
  .dashboard-grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:28px; }}
  .stat-card {{ background:#111; border:1px solid #1e1e1e; border-radius:6px; padding:16px; }}
  .stat-number {{ font-size:28px; font-weight:700; color:#fff; }}
  .stat-label {{ font-size:11px; color:#555; margin-top:4px; text-transform:uppercase; letter-spacing:0.5px; }}
  .doc-table {{ width:100%; border-collapse:collapse; background:#111; border-radius:6px; overflow:hidden; border:1px solid #1e1e1e; }}
  .doc-table th {{ background:#161616; padding:10px 14px; text-align:left; font-size:11px; color:#555; text-transform:uppercase; letter-spacing:0.8px; font-weight:600; border-bottom:1px solid #1e1e1e; }}
  .doc-table td {{ padding:10px 14px; font-size:13px; border-bottom:1px solid #161616; }}
  .doc-table tr:last-child td {{ border-bottom:none; }}
  .doc-table tr:hover td {{ background:#141414; cursor:pointer; }}
  .status-pill {{ display:inline-block; font-size:10px; padding:2px 8px; border-radius:12px; font-weight:600; }}
  .status-Live {{ background:#1a3a1a; color:#4caf50; }}
  .status-Draft {{ background:#3a2a1a; color:#ff9800; }}
  .status-Under-Review {{ background:#1a2a3a; color:#2196f3; }}
  .status-Expired {{ background:#3a1a1a; color:#f44336; }}
  .pub-yes {{ color:#4caf50; font-size:12px; }}
  .pub-no {{ color:#555; font-size:12px; }}
  .pub-pending {{ color:#ff9800; font-size:12px; }}
  .detail-panel {{ background:#111; border:1px solid #1e1e1e; border-radius:6px; padding:24px; }}
  .detail-title {{ font-size:20px; font-weight:700; color:#fff; margin-bottom:16px; }}
  .detail-meta {{ display:flex; gap:16px; flex-wrap:wrap; margin-bottom:24px; padding-bottom:16px; border-bottom:1px solid #1e1e1e; }}
  .meta-item {{ font-size:12px; color:#666; }}
  .meta-item strong {{ color:#aaa; display:block; margin-bottom:2px; }}
  .content-section {{ margin-bottom:20px; padding:16px; background:#0d0d0d; border-radius:4px; border:1px solid #1a1a1a; }}
  .content-section h4 {{ font-size:12px; color:#555; text-transform:uppercase; letter-spacing:0.8px; margin-bottom:10px; }}
  .content-text {{ font-size:14px; color:#ccc; line-height:1.6; }}
  .draft-text {{ font-size:14px; color:#ff9800; line-height:1.6; font-style:italic; }}
  .notes-text {{ font-size:12px; color:#666; margin-top:8px; padding:8px; background:#111; border-radius:3px; border-left:2px solid #333; }}
  .copy-btn {{ background:#1e1e1e; border:1px solid #2a2a2a; color:#888; padding:6px 14px; border-radius:4px; font-size:12px; cursor:pointer; margin-top:12px; }}
  .copy-btn:hover {{ background:#252525; color:#ccc; }}
  .filter-bar {{ display:flex; gap:8px; margin-bottom:16px; flex-wrap:wrap; }}
  .filter-btn {{ background:#1a1a1a; border:1px solid #2a2a2a; color:#666; padding:5px 12px; border-radius:4px; font-size:12px; cursor:pointer; }}
  .filter-btn:hover, .filter-btn.active {{ background:#222; color:#ccc; border-color:#444; }}
  .section-title {{ font-size:13px; font-weight:600; color:#888; margin-bottom:12px; text-transform:uppercase; letter-spacing:0.5px; }}
  .generated {{ font-size:11px; color:#333; padding:16px 0 4px; text-align:right; }}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-brand"><span>SumoSim</span><span class="dot">•</span>Compliance Hub</div>
  <div class="topbar-meta">Generated {generated} — edit SumoSim_Compliance_Hub.xlsx then re-run generate_hub.py</div>
</div>
<div class="layout">
  <div class="sidebar" id="sidebar">
    <div class="sidebar-section">Overview</div>
    <div class="sidebar-item active" onclick="showDashboard(this)">📊 Dashboard</div>
    <div class="sidebar-section" style="margin-top:12px">Documents</div>
  </div>
  <div class="main" id="main"></div>
</div>

<script>
const DOCS = {docs_json};
const CONTENT = {content_json};

let currentFilter = 'All';

function statusClass(s) {{
  return 'status-' + s.replace(' ', '-');
}}

function pubClass(p) {{
  if(p==='Yes') return 'pub-yes';
  if(p==='Pending') return 'pub-pending';
  return 'pub-no';
}}

function buildSidebar() {{
  const sidebar = document.getElementById('sidebar');
  DOCS.forEach((doc, i) => {{
    const el = document.createElement('div');
    el.className = 'sidebar-item';
    el.id = 'nav-' + i;
    const badge = doc.status === 'Live' ? 'badge-live' :
                  doc.status === 'Draft' ? 'badge-draft' : 'badge-review';
    const short = doc.status === 'Live' ? 'L' : doc.status === 'Draft' ? 'D' : 'R';
    el.innerHTML = doc.title + `<span class="sidebar-badge ${{badge}}">${{short}}</span>`;
    el.onclick = function() {{ showDoc(i, this); }};
    sidebar.appendChild(el);
  }});
}}

function showDashboard(el) {{
  document.querySelectorAll('.sidebar-item').forEach(i => i.classList.remove('active'));
  el.classList.add('active');
  const live = DOCS.filter(d=>d.status==='Live').length;
  const draft = DOCS.filter(d=>d.status==='Draft').length;
  const review = DOCS.filter(d=>d.status==='Under Review').length;
  const pub = DOCS.filter(d=>d.published==='Yes').length;

  let filtered = currentFilter === 'All' ? DOCS : DOCS.filter(d =>
    currentFilter === 'Live' ? d.status==='Live' :
    currentFilter === 'Draft' ? d.status==='Draft' :
    currentFilter === 'Review' ? d.status==='Under Review' :
    d.category === currentFilter
  );

  const categories = [...new Set(DOCS.map(d=>d.category))];
  const filterBtns = ['All','Live','Draft','Review',...categories].map(f =>
    `<button class="filter-btn ${{currentFilter===f?'active':''}}" onclick="setFilter('${{f}}')">${{f}}</button>`
  ).join('');

  const rows = filtered.map((doc,i) => `
    <tr onclick="showDoc(${{DOCS.indexOf(doc)}}, document.getElementById('nav-${{DOCS.indexOf(doc)}}'))">
      <td style="color:#fff;font-weight:500">${{doc.title}}</td>
      <td style="color:#666">${{doc.category}}</td>
      <td style="color:#555">v${{doc.version}}</td>
      <td><span class="status-pill ${{statusClass(doc.status)}}">${{doc.status}}</span></td>
      <td style="color:#666">${{doc.nextReview}}</td>
      <td><span class="${{pubClass(doc.published)}}">${{doc.published==='Yes'?'✓ Live':doc.published==='Pending'?'⏳ Pending':'— Unpublished'}}</span></td>
    </tr>`).join('');

  document.getElementById('main').innerHTML = `
    <div class="dashboard-grid">
      <div class="stat-card"><div class="stat-number">${{DOCS.length}}</div><div class="stat-label">Total Documents</div></div>
      <div class="stat-card"><div class="stat-number" style="color:#4caf50">${{live}}</div><div class="stat-label">Live</div></div>
      <div class="stat-card"><div class="stat-number" style="color:#ff9800">${{draft}}</div><div class="stat-label">Draft</div></div>
      <div class="stat-card"><div class="stat-number" style="color:#4caf50">${{pub}}</div><div class="stat-label">Published on Site</div></div>
    </div>
    <div class="section-title">All Compliance Documents</div>
    <div class="filter-bar">${{filterBtns}}</div>
    <table class="doc-table">
      <thead><tr>
        <th>Document</th><th>Category</th><th>Version</th><th>Status</th><th>Next Review</th><th>Website</th>
      </tr></thead>
      <tbody>${{rows}}</tbody>
    </table>
    <div class="generated">Data from SumoSim_Compliance_Hub.xlsx · Generated {generated}</div>
  `;
}}

function setFilter(f) {{
  currentFilter = f;
  showDashboard(document.querySelector('.sidebar-item'));
}}

function showDoc(i, navEl) {{
  document.querySelectorAll('.sidebar-item').forEach(el => el.classList.remove('active'));
  if(navEl) navEl.classList.add('active');
  const doc = DOCS[i];
  const sections = CONTENT[doc.title] || [];

  const sectionsHtml = sections.map(s => `
    <div class="content-section">
      <h4>${{s.section}}</h4>
      ${{s.live ? `<div class="content-text">${{s.live}}</div>` : ''}}
      ${{s.draft && s.draft !== s.live ? `<div class="draft-text">📝 Draft: ${{s.draft}}</div>` : ''}}
      ${{s.notes ? `<div class="notes-text">💡 ${{s.notes}}</div>` : ''}}
    </div>
  `).join('');

  const allText = sections.map(s => `${{s.section}}\\n${{s.live || s.draft}}`).join('\\n\\n');

  document.getElementById('main').innerHTML = `
    <div class="detail-panel">
      <div style="margin-bottom:16px">
        <span style="color:#555;font-size:12px;cursor:pointer" onclick="showDashboard(document.querySelector('.sidebar-item'))">← Dashboard</span>
      </div>
      <div class="detail-title">${{doc.title}}</div>
      <div class="detail-meta">
        <div class="meta-item"><strong>Category</strong>${{doc.category}}</div>
        <div class="meta-item"><strong>Version</strong>v${{doc.version}}</div>
        <div class="meta-item"><strong>Status</strong><span class="status-pill ${{statusClass(doc.status)}}" style="font-size:11px">${{doc.status}}</span></div>
        <div class="meta-item"><strong>Reviewed By</strong>${{doc.reviewedBy}}</div>
        <div class="meta-item"><strong>Last Reviewed</strong>${{doc.lastReviewed}}</div>
        <div class="meta-item"><strong>Next Review</strong>${{doc.nextReview}}</div>
        <div class="meta-item"><strong>Website Page</strong>${{doc.websitePage || '—'}}</div>
        <div class="meta-item"><strong>Published</strong>${{doc.published}}</div>
      </div>
      ${{sectionsHtml || '<div style="color:#444;font-style:italic">No content sections yet.</div>'}}
      <button class="copy-btn" onclick="navigator.clipboard.writeText(decodeURIComponent('${{encodeURIComponent(allText)}}'))">📋 Copy content for CMS</button>
    </div>
  `;
}}

buildSidebar();
showDashboard(document.querySelector('.sidebar-item'));
</script>
</body>
</html>"""
    return html


def main():
    print(f"Reading {XL_FILE}...")
    docs, content = load_data(XL_FILE)
    print(f"  Found {len(docs)} documents, {sum(len(v) for v in content.values())} content sections")
    html = generate_html(docs, content)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  Generated {OUTPUT_FILE} ({len(html)//1024}KB)")
    print(f"\nDone! Upload {OUTPUT_FILE} to GitHub repo to update the hub.")

if __name__ == '__main__':
    main()
