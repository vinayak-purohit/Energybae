# ⚡ EnergyBae — AI-Powered Solar Load Proposal Generator

## What This Project Does

Automates the manual process of reading a customer's electricity bill and generating a solar load proposal in Excel.

**Before:** Sales team spends 15–30 mins manually reading a bill and typing data into Excel.  
**After:** Upload the bill → AI reads it → Excel fills automatically → Download in seconds.

---

## How It Works

```
Bill Image / PDF (up to 5 files)
        ↓
  Groq Vision AI (LLaMA 4 Scout)
  reads it like a human reads a bill
        ↓
  Extracts: Name, Consumer No., Load,
  Connection Type, 12 months of usage
        ↓
  openpyxl fills those values into Excel
  (WITHOUT touching the formula cells)
        ↓
  Download filled Excel with solar
  calculations ready
```

---

## Tech Stack

| Technology | Purpose |
|---|---|
| **FastAPI** | Backend REST API server |
| **Groq Vision AI** | Reads electricity bills (LLaMA 4 Scout model) |
| **openpyxl** | Fills Excel template without breaking formulas |
| **PyMuPDF (fitz)** | Converts PDF bills to images for AI processing |
| **HTML / CSS / JS** | Professional frontend with EnergyBae branding |
| **Uvicorn** | ASGI server to run FastAPI |

---

## Project Structure

```
├── backend/
│   ├── app.py                  # FastAPI backend (AI extraction + Excel fill)
│   ├── requirements.txt        # Python dependencies
│   └── templates/
│       └── Energybae_Customer_Proposal.xlsx  # Bundled Excel template
│
├── frontend/
│   ├── index.html              # Main UI page
│   ├── style.css               # EnergyBae branded styles
│   ├── script.js               # Upload, API calls, download logic
│   └── assets/
│       ├── logo.svg            # EnergyBae logo
│       ├── hero_solar.png      # Hero section image
│       └── about_preview.png   # About section image
│
├── tests/
│   ├── api_pipeline_test.py    # Full API integration test
│   ├── test_e2e.py             # End-to-end pipeline test (standalone)
│   └── check_excel.py          # Excel formula verification
│
├── docs/
│   └── archive/                # Original bill images, old scripts, assignment PDF
│
├── Project_Walkthrough.md      # Development journey and decisions
└── README.md                   # This file
```

---

## How to Run

### Step 1: Install dependencies
```bash
cd backend
pip install -r requirements.txt
```

### Step 2: Start the server
```bash
cd backend
python -m uvicorn app:app --host 0.0.0.0 --port 8000
```

### Step 3: Open in browser
Go to **http://localhost:8000**

### Step 4: Use the app
1. Upload your electricity bill (PDF or image — up to 5 files)
2. Click **⚡ Generate Solar Proposal**
3. View the extracted data on screen
4. Click **Download Solar Proposal (Excel)**

---

## API Endpoint

### `POST /api/generate-proposal`

Upload one or more bill images/PDFs and receive a filled Excel proposal.

**Request:** `multipart/form-data` with field `files`  
**Response:** Excel file (`.xlsx`) with `Content-Disposition` header

```bash
# Example with curl
curl -X POST http://localhost:8000/api/generate-proposal \
  -F "files=@bill_page1.jpg" \
  -F "files=@bill_page2.jpg" \
  --output proposal.xlsx
```

---

## Key Technical Decisions

### Why Groq Vision AI (not Gemini)?
We initially used Google Gemini, but hit rate-limit errors during testing. Groq's API is faster and more reliable for our use case. We use the LLaMA 4 Scout model via the OpenAI-compatible client.

### Why multi-image input?
Single images sometimes gave inaccurate readings (especially for consumption tables). Sending multiple pages/angles in a single API call lets the AI cross-reference and return more accurate data.

### Why openpyxl (not pandas)?
The Excel template has formulas (like `=AVERAGE(D9:D21)`) that calculate the solar load automatically. Pandas would wipe all formulas when saving. openpyxl only writes to the specific data cells we target — leaving every formula intact.

### Why FastAPI (not Streamlit)?
We started with Streamlit for quick prototyping, but switched to FastAPI + a custom HTML/CSS/JS frontend for a professional, production-ready look inspired by the actual EnergyBae website.

---

## Excel Template Cell Mapping

### Data cells (filled by AI):

| Field | Cell(s) |
|---|---|
| Consumer Name | D1 |
| Consumer Number | D2 |
| Fixed Charges (Rs.) | D3 |
| Sanctioned Load (KW) | D4 |
| Connection Type | D5 |
| Monthly Month dates | C9 to C20 |
| Monthly Units consumed | D9 to D20 |
| Monthly Bill Amount | E9 to E20 |

### Formula cells (NEVER touched):

| Cell | Formula | Purpose |
|---|---|---|
| D22 | `=AVERAGE(D9:D21)` | Average monthly units |
| D23 | `=(D22*12*1.1)/1400` | Required kW capacity |
| D24 | `=D23/$C$7*1000` | Number of solar panels |
| D25 | `=ROUND(D24,0)*$C$7/1000` | Solar capacity (kW) |
| D26 | `=D25/$C$7*1000` | Number of panels (rounded) |
| D29 | `=SUM(25:25)` | Total solar capacity |
| D30 | `=SUM(26:26)` | Total number of panels |

---

## Running Tests

```bash
# Start the server first, then:

# Full API pipeline test (uploads real bills, checks response)
python tests/api_pipeline_test.py

# Standalone E2E test (Groq + Excel, no server needed)
python tests/test_e2e.py

# Excel formula integrity check
python tests/check_excel.py
```

---

## What I Would Improve Next
1. Pre-process bill images (crop important sections) before sending to AI for higher accuracy
2. Support for multiple consumers in one run
3. Auto-detect cell positions from any Excel template
4. Generate a PDF proposal alongside the Excel output
5. Add WhatsApp integration to receive bills directly
