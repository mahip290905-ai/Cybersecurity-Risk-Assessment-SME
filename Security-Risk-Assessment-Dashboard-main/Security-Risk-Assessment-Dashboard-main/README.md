# Security Risk Assessment Dashboard (NIST CSF 2.0 / ISO 27001)

A professional-grade cybersecurity risk assessment dashboard for Small and Medium Enterprises (SMEs). Aligned with the **NIST Cybersecurity Framework (CSF) 2.0** and **ISO/IEC 27001**, enabling organizations to evaluate, quantify, and visualize cyber risks effectively.

---

## Project Overview

- **SME Cyber Risk Research:** Built-in knowledge base covering 6 common small business threats (Phishing, Ransomware, Weak Passwords, Malware, Data Breaches, Insider Threats).
- **NIST CSF 2.0 Alignment:** Comprehensive coverage across all six functions: **Govern, Identify, Protect, Detect, Respond, Recover**.
- **Quantitative Evaluation Metric:** Strict 5x5 Likelihood x Impact scoring (1-25 scale) with Low, Medium, High, and Critical severity bands.
- **Interactive Visualizations:** 6-axis NIST CSF radar charts and 5x5 color-coded risk heatmaps.
- **Automated Mitigation Engine:** Automatically pairs identified threats with practical technical safeguards.
- **Simulated Case Studies:** Validated against three SME profiles (Retail Shop, Healthcare Clinic, Small IT Company).
- **Audit Reports:** One-click single-page executive PDF security report export via ReportLab and JSON export.
- **Secure Authentication:** User registration with password confirmation and stateless JWT tokens.

---

## Single-Command Quick Start

You can launch all services (**MongoDB + FastAPI Backend + React Frontend**) with **one single command** from the project root:

### In Windows Command Prompt (CMD):
```cmd
run
```
*(or double-click `run.bat`)*

### To Stop All Services:
```cmd
stop
```
*(or double-click `stop.bat`)*

### Alternative CLI Launchers:
```bash
# Using Python:
python run.py

# Using npm:
npm start
```

---

## Clean Project Structure

```
Security-Risk-Assessment-Dashboard/
│
├── backend/                  # Python FastAPI REST API
│   ├── app/
│   │   ├── api/v1/           # API routes (auth, assessments, export)
│   │   ├── core/             # Configuration & JWT security
│   │   ├── db/               # MongoDB Motor connection & Pydantic schemas
│   │   ├── services/         # Risk model & NIST CSF 2.0 scoring logic
│   │   └── main.py           # FastAPI application entry point
│   ├── data/db/              # Local MongoDB database storage
│   ├── venv/                 # Python virtual environment
│   └── requirements.txt      # Python dependencies
│
├── frontend/                 # React.js SPA Application
│   ├── public/               # Static assets & HTML index
│   ├── src/
│   │   ├── components/       # Reusable UI widgets (Navbar, RadarChart, HeatMap)
│   │   ├── data/             # Threats research knowledge base (threats.js)
│   │   ├── pages/            # App pages (Dashboard, SMEResearch, NISTFramework, etc.)
│   │   └── App.js            # Router switchboard
│   └── package.json          # Node.js dependencies & scripts
│
├── report/                   # Final Academic Project Deliverables
│   ├── Minor_Project_1_Report.docx   # Word Report (A4, Times New Roman, Double-bordered)
│   ├── Minor_Project_1_Report.pdf    # Compiled PDF Report
│   ├── report_notes.md               # Technical notes & verification mapping
│   ├── build_report.py               # Document builder script
│   └── downloaded_test_preview.png   # PDF executive report preview figure
│
├── .env                      # Environment variables
├── .gitignore                # Git ignore patterns
├── LICENSE                   # MIT Open Source License
├── package.json              # Root npm launcher scripts
├── README.md                 # Project documentation
├── run.bat                   # 1-Click Windows CMD service launcher
├── run.py                    # Cross-platform Python service launcher
├── stop.bat                  # 1-Click Windows CMD service stopper
└── stop.py                   # Cross-platform Python service stopper
```

---

## Service Endpoints & Ports

| Service | Port / URL | Purpose |
|---|---|---|
| **React Frontend** | `http://localhost:3000` | Dark-theme user interface dashboard |
| **FastAPI Backend** | `http://127.0.0.1:8000` | REST API gateway & risk engine |
| **API Documentation** | `http://127.0.0.1:8000/docs` | Swagger interactive API docs |
| **MongoDB Server** | `mongodb://127.0.0.1:27017` | Local NoSQL document store |

---

## License
This project is licensed under the MIT License.
