# 📊 Attitude Campaign Report Generator

Internal tool to automatically generate posting reports from Google Drive campaign folders.

**Built for**: Attitude Ideology Sdn Bhd team
**Stack**: Python + Streamlit + python-pptx + Tesseract OCR

---

## ✨ What It Does

Team members paste a Google Drive campaign folder link, fill in 3-4 fields, and get a polished
PowerPoint report in 30 seconds — no manual data entry, no copy-paste from sheets.

```
Drive folder → CSV from response sheet → Download screenshots →
OCR-pick correct insight image → Fill template.pptx → Download .pptx
```

---

## 🚀 Quick Start (For Team Users)

1. Open the report tool URL (shared by your manager).
2. Make sure your campaign Drive folder is shared as **"Anyone with link can view"**.
3. Paste the folder link, fill in:
   - Campaign Name (e.g. "Jombeli Ampang Point")
   - Campaign Month (e.g. "Apr 2026")
   - Client (Watsons is default)
   - Hero image (optional)
4. Click **Generate Report**.
5. If asked, confirm which insight screenshot is correct for any KOC the system isn't sure about.
6. Download your `.pptx` and use as-is, or polish further in PowerPoint.

---

## 📁 Project Structure

```
attitude_report_app/
├── app.py                  # Streamlit main app
├── drive_reader.py         # Reads Drive folder + parses response sheet
├── image_downloader.py     # Downloads all screenshots from Drive
├── ocr_classifier.py       # OCR-picks correct insight screenshot
├── color_extractor.py      # Extracts dominant color from hero image
├── pptx_builder.py         # Fills template.pptx with campaign data
├── template.pptx           # The slide template (editable in PowerPoint)
├── brands/                 # Client logos (PNG/JPG)
│   └── watsons.png
├── requirements.txt        # Python packages
├── packages.txt            # System packages (tesseract for OCR)
├── README.md               # This file
├── SETUP.md                # Deployment guide
└── TEMPLATE_GUIDE.md       # How to edit template.pptx (for designers)
```

---

## 🔧 Requirements

The campaign Drive folder must follow this structure:

```
Campaign Folder/
├── XHS/
│   └── [Form Responses Sheet]  ← read by tool
└── TIKTOK/
    └── [Form Responses Sheet]  ← read by tool
```

Each response sheet must have these columns (in this order):
```
Timestamp | Email | Creator Name | Posted Link | Post Screenshot | Insight Screenshot |
Views | Likes | Comments | Saves | Shares | Impressions |
IG Profile Name | IG Posted Link | IG Post Screenshot | IG Insight Screenshot |
Views | Likes | Comments | Saves | Shares | Invoice
```

---

## 📚 Documentation

- **[SETUP.md](SETUP.md)** — How to install and deploy the tool
- **[TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md)** — How to edit `template.pptx` (for designers)

---

## 🐛 Troubleshooting

**"Cannot access folder"**
→ The Drive folder isn't shared. Set to "Anyone with the link can view".

**"No campaign data found"**
→ Either subfolder names aren't matching ("XHS" / "TIKTOK"), or response sheets aren't there yet.

**Screenshots don't appear in the PPT**
→ The KOC's Drive permissions on individual screenshots may not be public. The tool will show
  placeholders — drop in the images manually in PowerPoint.

**Wrong insight screenshot picked**
→ During generation, if the system isn't sure, it'll ask you to pick. If it picked wrong silently,
  you can swap the image in PowerPoint manually (right-click → Change Picture).

---

## 📝 License

Internal tool — not for redistribution.
