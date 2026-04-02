<div align="center">

# 🦟 Dengue: Clinical Intelligence Dashboard

> A real-time clinical research dashboard for dengue surveillance at **Colombo South Teaching Hospital, Sri Lanka**

<br/>

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=for-the-badge&logo=chartdotjs&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google_Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)

<br/>

![Status](https://img.shields.io/badge/Status-Active_Enrolment-brightgreen?style=flat-square)
![Study](https://img.shields.io/badge/Study-Single_Centre_Prospective-blue?style=flat-square)
![Variables](https://img.shields.io/badge/Variables_Assessed-323-purple?style=flat-square)
![Institution](https://img.shields.io/badge/Institution-CSTH_Sri_Lanka-red?style=flat-square)

</div>

---

## 🔍 Overview

This dashboard provides clinicians and researchers a single interface to monitor patient enrolment, assess severity, track data quality, and explore clinical patterns — all updated in real time from the hospital's Google Forms data collection system.

| Question | Where to look |
|---|---|
| **What is happening?** | Dashboard → KPIs, enrollment trend, data completeness |
| **Who is at risk?** | Severity × Risk Matrix → high-risk patient list |
| **How is data quality?** | Collection Monitor → missing data, outlier flags |

---

## 🌐 Live Dashboard

Deployed via Google Apps Script Web App — access requires Google account authorization.

---

## ⚡ Features

### 🧠 Clinical Intelligence
- **Severity × Risk Cross-Analysis** — binary WHO severity staging cross-tabulated with cumulative risk scores, surfacing "silent danger" patients with mild presentation but high/critical risk
- **High-Risk Patient Table** — ranked list with individual risk factors (age, comorbidities, medications, bleeding)
- **WHO Warning Signs Panel** — frequency analysis of all 8 warning sign criteria

### 📡 Real-Time Monitoring
- **Collection Monitor Tab** — tracks enrolment progress, completeness over time, follow-up phase completion (5 phases), and days-to-completion distributions
- **Fix Today List** — patient IDs with incomplete data, updated on every load
- **Adjusted Completeness Score** — excludes structurally absent fields (conditional sub-fields, missing by design)

### 🗂️ Dashboard Tabs

| Tab | Content |
|---|---|
| 📊 Dashboard | KPIs, severity × risk matrix, warning signs, data quality |
| 👥 Demographics | Age groups, gender, geographic distribution |
| 🩺 Clinical | Symptom frequency, vital signs, warning signs |
| 🏷️ WHO Classification | DF / DHF / DSS distribution |
| 🌿 Lifestyle | Comorbidities, risk factors, medications |
| 🔬 Outcomes & Labs | ICU/HDU, hospital stay, platelet nadir, serology, complications |
| 📋 Collection Monitor | Enrolment trend, completeness, follow-up tracker, patient gap list |

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Google Apps Script (JavaScript) |
| **Data source** | Google Sheets (Google Forms responses) |
| **Frontend** | HTML5 + CSS3 + vanilla JavaScript |
| **Charts** | Chart.js |
| **Hosting** | Google Apps Script Web App (`HtmlService`) |
| **Caching** | `CacheService.getScriptCache()` — 300s main, 60s fresh monitor |

---

## 📁 Repository Structure

```
dengue-dashboard/
├── Code.gs               # Backend — data fetch, column mapping, all calculations
├── Index.html            # Frontend — single-page dashboard (HTML + CSS + JS)
├── appsscript.json       # GAS project configuration
└── README.md
```

---

## 🚀 Deployment

### 1. Create a Google Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Copy the contents of `Code.gs` into the default `Code.gs` file
3. Create a new HTML file named `Index` and paste the contents of `Index.html`
4. Copy `appsscript.json` into the project manifest (enable **Show manifest file** in Project Settings)

### 2. Point to your Google Sheet

In `Code.gs`, update:

```javascript
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const SHEET_NAME     = 'Form Responses 1';
```

The Spreadsheet ID is in the sheet URL:
`docs.google.com/spreadsheets/d/[ID_HERE]/edit`

### 3. Deploy as a Web App

1. **Deploy** → **New deployment** → type: **Web app**
2. Set **Execute as**: `Me`
3. Set **Who has access**: `Anyone` (or restrict as needed)
4. Click **Deploy** → authorize permissions → copy the web app URL

### 4. Refresh data

- The dashboard caches data for 5 minutes. Use the **Refresh** button to force a reload.
- Collection Monitor has a separate 60-second cache for near-real-time updates.

---

## ⚙️ Backend Architecture

```
doGet()
 └── getDashboardData()
      ├── getColumnIndices(headers)        — maps ~40+ form fields to column indices
      ├── calculateLabValues(data, idx)    — platelet, WBC, ALT/AST, creatinine, albumin, INR
      ├── calculateSerology(data, idx)     — NS1, IgG/IgM, primary vs secondary
      ├── calculateOutcomes(data, idx)     — ICU/HDU, hospital days, diagnosis, complications
      ├── calculateWarningSigns(data, idx) — 8 WHO warning sign criteria
      ├── calculateRiskAssessment(data)    — cumulative risk score → Low/Moderate/High/Critical
      └── calculateCollectionMonitor()    — completeness, follow-up phases, enrollment trend
```

---

## 🔒 Security Notes

- This dashboard is intended for authorized research personnel only
- Restrict web app access appropriately before sharing externally
- Google Sheets permissions control who can modify the underlying data

---

## 📄 License

MIT License — free to use, fork, and modify for personal or research projects.

---

<div align="center">

*Built for clinical research — not for diagnostic use.*

---

**Built by Dilshan Ganepola** — Principal AI Specialist

MRCGP &nbsp;|&nbsp; MSc in Health Informatics &nbsp;|&nbsp; US Patent in Medical AI

*Strong interest in precision AI*

</div>
