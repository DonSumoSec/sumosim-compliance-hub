# SumoSim Compliance Hub

> Internal compliance document management system for SumoSim Ltd.  
> Single source of truth: `SumoSim_Compliance_Hub.xlsx` → auto-builds → deploys live.

**Live Hub:** https://donsumosec.github.io/sumosim-compliance-hub/

---

## How It Works

```
Edit XL in Excel
    → Save to this cloned folder
        → GitHub Desktop: Commit + Push
            → GitHub Actions runs generate_hub.py (~60 seconds)
                → index.html rebuilt and deployed
                    → Live hub updates automatically
```

---

## Updating the Compliance Hub (Current Workflow)

### Step 1 — Edit the spreadsheet
Open `SumoSim_Compliance_Hub.xlsx` from this cloned folder in Excel.  
Make your changes across any sheet (Dashboard, Content Library, etc.).  
**Save the file back to this same folder.**

### Step 2 — Open GitHub Desktop
Switch to the `sumosim-compliance-hub` repo.  
You'll see `SumoSim_Compliance_Hub.xlsx` listed as a changed file.

### Step 3 — Commit
In the bottom-left Summary box, type a short description:
```
Update compliance docs - [what you changed]
```
Click **Commit to main**.

### Step 4 — Push
Click **Push origin** in the top bar.

### Step 5 — Wait ~60 seconds
The hub rebuilds automatically. Check live at:  
https://donsumosec.github.io/sumosim-compliance-hub/

---

## Repo Structure

```
sumosim-compliance-hub/
├── .github/
│   └── workflows/
│       └── build.yml                ← GitHub Actions pipeline (do not edit)
├── SumoSim_Compliance_Hub.xlsx      ← MASTER SOURCE — edit this
├── generate_hub.py                  ← Reads XL, generates index.html
├── index.html                       ← Auto-generated — do not edit manually
└── README.md                        ← This file
```

---

## Repo Files Explained

| File | Purpose | Edit? |
|---|---|---|
| `SumoSim_Compliance_Hub.xlsx` | Master compliance database | ✅ Yes — this is your data |
| `generate_hub.py` | Python script that builds the HTML hub | ⚠️ Only if adding new features |
| `index.html` | The live compliance hub webpage | ❌ Never — auto-generated |
| `.github/workflows/build.yml` | CI/CD pipeline | ❌ Never |

---

## Spreadsheet Structure

The XL has 5 sheets:

| Sheet | Purpose |
|---|---|
| 📊 Dashboard | 14 compliance documents — status, versions, review dates |
| 📝 Content Library | Copyable text sections for each document |
| 📋 Version History | Change log for all documents |
| 🌐 Website Update Workflow | Checklist for publishing docs to website |
| 📅 Review Calendar | Upcoming review schedule |

---

## SharePoint Integration

The master XL lives in SharePoint at:  
`https://sumosim.sharepoint.com/sites/SumosimLtd` → Compliance Source Files

The hub is embedded in SharePoint at:  
`https://sumosim.sharepoint.com/sites/SumosimLtd/SitePages/Compliance-Hub.aspx`

---

## 14 Compliance Documents

| Document | Category | Status |
|---|---|---|
| Terms and Conditions v1.4 | Legal | Under Review |
| Privacy Policy v2.2 | Privacy | Under Review |
| Cookie Policy v1.1 | Privacy | Under Review |
| Acceptable Use Policy v1.1 | Legal | Live |
| Vulnerability Situation Policy v1.0 | Security | Under Review |
| Privacy Notice v1.2 | Privacy | Live |
| Security Policy v1.0 | Security | Draft |
| Customer Code of Practice v1.0 | Consumer | Live |
| Accessibility Policy v1.0 | Consumer | Live |
| MVNO Fair Use Policy v1.0 | Telecoms | Draft |
| Crypto Rewards T&Cs v1.1 | Crypto | Draft |
| Data Retention Policy v1.0 | Privacy | Draft |
| Complaints Policy v1.1 | Consumer | Live |
| Gamification & Referral Policy v1.1 | Legal | Draft |

---

## Planned: Phase 2 — MCP Server Integration (Post Pre-Seed)

> **Roadmap item — to be built after pre-seed funding secured.**

The current GitHub Desktop → manual push workflow will be replaced with a  
fully automated **Model Context Protocol (MCP) server** that connects the  
compliance hub directly to Claude AI, SharePoint, and internal SumoSim tooling.

### What the MCP Server Will Do

```
SharePoint XL saved (auto-detected)
    → MCP server pushes XL to GitHub via API
        → GitHub Actions rebuilds hub (~60s)
            → Claude AI can query/update compliance docs via chat
                → Zero manual steps end-to-end
```

### Planned MCP Capabilities

| Tool | Description |
|---|---|
| `get_document_status` | Query any compliance doc's current status, version, review date |
| `update_document_status` | Update status, reviewer, dates directly from Claude chat |
| `get_review_schedule` | List all documents due for review in next N days |
| `publish_document` | Mark a document as published and trigger hub rebuild |
| `add_content_section` | Add or update a content section in the Content Library |
| `trigger_rebuild` | Force a hub rebuild without changing the XL |
| `sync_from_sharepoint` | Pull latest XL from SharePoint and push to GitHub |
| `compliance_summary` | Return a full status report of all 14 documents |

### Planned Architecture

```
Claude (AI layer)
    ↕ MCP protocol
SumoSim Compliance MCP Server (Node.js / Python)
    ↕ GitHub API         → sumosim-compliance-hub repo
    ↕ SharePoint API     → Compliance Source Files library
    ↕ Microsoft Graph    → Power Automate triggers
    ↕ SendGrid / SMTP    → Review reminder emails
```

### MCP Server Stack (Planned)

- **Runtime:** Node.js (TypeScript) or Python
- **MCP SDK:** `@anthropic-ai/mcp` or `mcp` Python package
- **GitHub integration:** Octokit REST API (handles SHA automatically)
- **SharePoint integration:** Microsoft Graph API + MSAL auth
- **Hosting:** Azure Functions or a lightweight VPS (post-seed)
- **Auth:** GitHub PAT + Azure App Registration (service principal)

### Why MCP (vs Power Automate)

Power Automate handles the file sync trigger but has no AI layer —  
it can push a file but can't understand, query, or reason about the compliance data.

An MCP server gives Claude direct read/write access to the compliance system,  
enabling workflows like:

> *"Which documents are due for review in the next 30 days?"*  
> *"Mark the Privacy Policy as Live and trigger a rebuild."*  
> *"Draft a review summary for the Gamification Policy and email it to the legal team."*

This turns the compliance hub from a static dashboard into an  
**AI-native compliance management system** — a core part of SumoSim's  
lean, automated internal ops stack.

### Development Timeline (Estimated)

| Milestone | Target |
|---|---|
| Pre-seed closed | Q2 2026 |
| MCP server MVP (read-only tools) | 4 weeks post-seed |
| Write tools + SharePoint sync | 6 weeks post-seed |
| Full automation — zero manual steps | 8 weeks post-seed |

---

## Company Info

| | |
|---|---|
| **Company** | SumoSim Ltd |
| **Companies House** | 16233377 |
| **ICO Registration** | ZB980254 |
| **Registered Address** | 71-75 Shelton Street, Covent Garden, London, WC2H 9JQ |

---

*Last updated: March 2026*
