# CMMC L2 SSP Drafting Tool

A **100% local** web application for drafting CMMC Level 2 System Security Plans (SSPs).
No data is sent to any cloud service. All data is stored in the `data/` folder on your device.

---

## Quick Start

### Windows
1. Double-click **`start.bat`**
2. Open your browser to **http://127.0.0.1:5000**

### macOS / Linux
```bash
chmod +x start.sh
./start.sh
```
Then open your browser to **http://127.0.0.1:5000**

### Manual Start
```bash
pip install -r requirements.txt
python app.py
```

---

## Demo Hosting (Render)

This tool is built to run **100% locally** — that's the intended way to use it for real SSP work,
since CMMC/CUI-adjacent data shouldn't touch a third-party server. But for a quick demo or walkthrough,
it can be deployed to [Render](https://render.com) as a standard Python web service.

**Important:** Render's filesystem is ephemeral. Anything saved (SSP data, uploaded images) will be
**wiped on every redeploy or restart**. This is fine for click-through demos, but don't rely on it for
anything you need to keep — use the local app for actual drafting work.

### Deploy via Render Blueprint (recommended)
1. Push this folder to a GitHub repo.
2. In the Render dashboard, choose **New → Blueprint**, point it at the repo. `render.yaml` configures
   everything automatically (build command, start command, free plan).
3. Render builds and deploys; you'll get a `https://<your-app>.onrender.com` URL.

### Deploy manually
1. Push this folder to a GitHub repo.
2. In Render, choose **New → Web Service**, connect the repo.
3. Set:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 120`
   - **Plan:** Free (or Starter for no spin-down delay)
4. Deploy.

The free plan spins the service down after 15 minutes of inactivity — the first request after that
takes 30-60 seconds to wake back up. For a live demo session, an upgraded plan removes the cold start.

---

## Features

### SSP Drafting
- **System Information** — Organization details, contacts, system identification
- **Environment Description** — Narrative sections for system, physical, network, cloud environments
- **Diagrams & Images** — Upload and embed logo, network diagram, physical site diagram into SSP
- **CMMC Controls Status** — Mark all 110 NIST SP 800-171 Rev. 2 controls as Implemented / Planned / N/A / Not Implemented

### Asset Inventory (CMMC Scoping Guide L2 v2.13)
- **IT Asset Tracker** — All network/compute/cloud assets with CMMC category assignment
- **Physical Security Asset Tracker** — Badge readers, cameras, alarm panels, biometric devices, etc.
- Import from an existing CMMC L2 Artifacts Tracker Excel file
- Export to formatted Excel workbook with color-coded categories

### Asset Categories (per CMMC Scoping Guide L2 v2.13)
| Code | Category | Scope |
|------|----------|-------|
| **CUI** | CUI Asset | ✅ In scope — all 110 controls |
| **SPA** | Security Protection Asset | ✅ In scope — relevant controls |
| **CRMA** | Contractor Risk Managed Asset | ✅ In scope — limited assessment |
| **SA-GFE/IoT/OT/RI/TE** | Specialized Assets | ✅ In scope — document in SSP |
| **OOS** | Out-of-Scope Asset | ❌ Not in scope |

### Generated Outputs
1. **SSP.docx** — Full System Security Plan document
   - Cover page with organization logo
   - Section 1: System Identification (all subsections)
   - Physical Security Environment with embedded physical diagram
   - Networking section with embedded network diagram
   - Server and asset inventory tables
   - Section 2: Controls summary table
   - Section 3: Full CMMC compliance section (all 110 controls)
   - Version history table

2. **Asset Tracker.xlsx** — Multi-sheet Excel workbook
   - Sheet 1: Full IT Asset Inventory with CMMC categories
   - Sheet 2: Physical Security Asset Inventory
   - Sheet 3: Authorized Users
   - Sheet 4: Authorized Server Apps
   - Sheet 5: Authorized System Processes
   - Sheet 6: Authorized Ports & Protocols
   - Sheet 7: Category Reference guide

---

## Data Storage
All data is stored locally in:
- `data/ssp_data.json` — SSP content
- `data/asset_data.json` — Asset inventory
- `uploads/` — Uploaded images (logo, diagrams)

**No internet connection required after initial setup.**

---

## SSP Template Reference
Standards: NIST SP 800-171 Rev. 2 / CMMC Level 2
Scoping: CMMC Scoping Guide Level 2 v2.13 (DoD-CIO-00006)
