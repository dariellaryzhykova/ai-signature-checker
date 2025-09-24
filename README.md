# Signature Checker (Google Apps Script + Document AI)

An AI **workflow automation** that scans Google Drive for PDFs(OCR formats), sends them to **Google Document AI**, extracts tables, and verifies whether **Signatures** are present. Results can be emailed and/or posted to Slack.

> This is an AI-powered **workflow**, not a fully autonomous agent. It follows your rules, reliably.

---

## ✨ What it does
- Watches Drive folders for new PDFs (e.g., files ending with `RTR.pdf`)
- Processes each file with **Document AI Form Parser**
- Extracts rows for the **Quality/QA** function and checks for signature presence
- Sends a **summary email** (and optionally **Slack** message)
- Can run on a **5-minute trigger** automatically

---

## 🧩 Architecture
- **Google Drive**
- **Google Apps Script** → orchestrates scan + calls Document AI + outputs
- **Document AI (Form Parser)** → parses tables and text
- **Gmail / Slack** → notifications

---

## ✅ Compliance & Certifications (Document AI service)
- ISO 27001, 27017, 27018  
- SOC 2, SOC 3  
- PCI DSS  
- FedRAMP High (U.S.)  
- HIPAA (healthcare)

> These are for **Google Cloud Document AI**. Review Google’s latest documentation for current status.

---

## 🚀 Quick Start

### 1) Google Cloud setup (one-time)
1. Create or choose a project and enable **Billing**  
2. Enable APIs:
   - **Document AI API**
   - **Drive API** (for REST fallback)
3. Create a **Document AI processor** (Location: `us`, Type: *Form Parser*)  

### 2) Link Apps Script to the Cloud project
- Apps Script → **Project Settings** → **Google Cloud Platform (GCP) Project** → **Change project** → paste **Project Number**.

### 3) Apps Script `appsscript.json`
```json
{
  "timeZone": "America/Chicago",
  "runtimeVersion": "V8",
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/cloud-platform",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/gmail.send"
  ]
}
