# Dengue: Clinical Intelligence Dashboard

> A real-time clinical research dashboard for dengue multi-variable study at **Colombo South Teaching Hospital (CSTH)**, Sri Lanka — built on Google Apps Script and powered by live Google Sheets data.

---

## Overview

This dashboard provides clinicians and researchers a single interface to monitor patient enrolment, assess severity, track data quality, and explore clinical patterns — all updated in real time from the hospital's Google Forms data collection system.

The system answers three core clinical questions:

| Question | Where to look |
|---|---|
| **What is happening?** | Dashboard → KPIs, enrollment trend, data completeness |
| **Who is at risk?** | Severity × Risk Matrix → high-risk patient list |
| **How is data quality?** | Collection Monitor → missing data, outliers, issues |

---

## Live Dashboard

Deployed via Google Apps Script Web App — access requires Google account authorization.

---

## Features

### Clinical Intelligence
- **Severity × Risk Cross-Analysis** — binary WHO severity staging cross-tabulated with cumulative risk scores, surfacing "silent danger" patients with mild presentation but high/critical risk
- **High-Risk Patient Table** — ranked list with individual risk factors (age, comorbidities, medications, bleeding)
- **WHO Warning Signs Panel** — frequency analysis of all 8 warning sign criteria

### Real-Time Monitoring
- **Collection Monitor Tab** — tracks enrolment progress, completeness over time, follow-up phase completion (5 phases), and days-to-completion distributions
- **Fix Today List** — patient IDs with incomplete data, updated on every load
- **Adjusted Completeness Score** — excludes structurally absent fields (conditional sub-fields, missing by design)

### Clinical Tabs
| Tab | Content |
|---|---|
| **Dashboard** | KPIs, severity×risk matrix, warning signs, data quality |
| **Demographics** | Age groups, gender, geographic distribution |
| **Clinical** | Symptom frequency, vital signs, warning signs |
| **WHO Classification** | DF / DHF / DSS distribution |
| **Lifestyle** | Comorbidities, risk factors, medications |
| **Outcomes & Labs** | ICU/HDU, hospital stay, platelet nadir, serology, complications |
| **Collection Monitor** | Enrolment trend, completeness, follow-up tracker, patient gap list |

### Data Quality Analysis (EDA Integration — in progress)
- Missing data by key variable (albumin, BMI, discharge platelet, liver enzymes)
- IQR-based outlier detection with clinical threshold flags
- Structured issue log (critical / moderate / minor / low severity)

---

## Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Google Apps Script (JavaScript) |
| **Data source** | Google Sheets (Google Forms responses) |
| **Frontend** | HTML5 + CSS3 + vanilla JavaScript |
| **Charts** | Chart.js |
| **Hosting** | Google Apps Script Web App (`HtmlService`) |
| **Caching** | `CacheService.getScriptCache()` — 300s main, 60s fresh monitor |

---

## Repository Structure

```
dengue-dashboard/
├── Code.gs               # Backend — data fetch, column mapping, all calculations
├── Index.html            # Frontend — single-page dashboard (HTML + CSS + JS)
├── appsscript.json       # GAS project configuration
├── data emelemnts.txt    # Column header reference for Google Forms (~300+ fields)
├── Dengue_MVS_EDA_Report.html  # EDA report (source for Data Analysis tab)
└── README.md
```

---

## Deployment

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
`docs.google.com/spreadsheets/d/**[ID_HERE]**/edit`

### 3. Deploy as a Web App

1. **Deploy** → **New deployment** → type: **Web app**
2. Set **Execute as**: `Me`
3. Set **Who has access**: `Anyone` (or restrict as needed)
4. Click **Deploy** → authorize permissions → copy the web app URL

### 4. Refresh data

- The dashboard caches data for 5 minutes. Use the **Refresh** button to force a reload.
- Collection Monitor has a separate 60-second cache for near-real-time updates.

---

## Backend Architecture

```
doGet()
 └── getDashboardData()
      ├── getColumnIndices(headers)       — maps ~40+ form fields to column indices
      ├── calculateLabValues(data, idx)   — platelet, WBC, ALT/AST, creatinine, albumin, INR
      ├── calculateSerology(data, idx)    — NS1, IgG/IgM, primary vs secondary
      ├── calculateOutcomes(data, idx)    — ICU/HDU, hospital days, diagnosis, complications
      ├── calculateWarningSigns(data, idx)— 8 WHO warning sign criteria
      ├── calculateRiskAssessment(data)   — cumulative risk score → Low/Moderate/High/Critical
      └── calculateCollectionMonitor()   — completeness, follow-up phases, enrollment trend
```

---

## Data Model

The study collects **323 variables** per patient via Google Forms. Key domains:

| Domain | Variables |
|---|---|
| Demographics | Age, gender, BMI, education |
| Comorbidities | HT, DM, IHD, asthma/COPD, CKD, CLD |
| Symptoms | Fever, headache, myalgia, vomiting, rash, bleeding (×8 types) |
| WHO warning signs | Abdominal pain, persistent vomiting, fluid accumulation, mucosal bleeding, lethargy, liver enlargement, elevated HCT with low platelet, rapid clinical deterioration |
| Lab values | Platelet (admission/nadir/discharge), WBC, ALT, AST, creatinine, albumin, HCT, INR |
| Serology | NS1 Ag, IgM, IgG |
| Outcomes | ICU/HDU, hospital stay days, discharge outcome, complications (ALF, AKI, encephalopathy) |

---

## Known Data Quality Issues

| Severity | Issue |
|---|---|
| Critical | 66% missing — Serum Albumin |
| Moderate | 28% missing — BMI |
| Moderate | Free-text fields not machine-readable (medication names, other symptoms) |
| Moderate | 25 unnamed survey artefact columns (Unnamed: 298–322) |
| Minor | Extreme outlier BMI=80 — likely data entry error |
| Minor | Conditional sub-fields missing by design (pregnancy, bleeding sub-fields) |
| Low | PII present — phone number column |

---

## Security Notes

- This dashboard is intended for authorized research personnel only
- Restrict web app access appropriately before sharing externally
- Phone numbers are present in source data — do not expose in public deployments
- Google Sheets permissions control who can modify the underlying data

---

## Study Context

**Institution:** Colombo South Teaching Hospital (CSTH), Sri Lanka  
**Study:** Dengue Multi-Variable Study (MVS) — prospective, single-centre  
**Design:** Ongoing enrolment; structured data collection via Google Forms  
**Variables assessed:** 323 per patient  
**Current cohort:** N = 50 (snapshot: Oct 15 – Nov 26, 2025)

---

*Built for clinical research — not for diagnostic use.*
