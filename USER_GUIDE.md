# Triple Duty Bond Generator — User Guide

## Introduction

The **Triple Duty Bond Generator** automatically creates Triple Duty Bond Word documents by extracting data from Bill of Entry (BE) PDF files. It reads the PDF, pulls out key fields (BE Number, Date, Duty Amounts, Packages, etc.), calculates the bond amount, and populates a pre-formatted Word template — preserving all bold/italic formatting in the final output.

**Who is this for?**
Customs operations staff and logistics teams at Nagarkot Forwarders who need to generate Triple Duty Bond documents quickly and accurately from Bill of Entry PDFs issued by Indian Customs (ICEGATE).

**Key Features:**
- One-click bond document generation from any Bill of Entry PDF.
- Automatic extraction of BE Number, Date, Importer, Duty Amount, Assessed Value, and Package count.
- Bond amount rounded up to the nearest ₹5,000 and converted to words (Indian numbering system).
- All original Word template formatting (bold, italic, fonts) is preserved.
- Output file is auto-named using the BE Number for easy identification.
- Bundled as a standalone `.exe` — no Python installation required.

---

## How to Use

### 1. Launching the App

Double-click the **`Triple_Duty_Bond_Generator.exe`** file located in the `dist/` folder.

The application will open in full-screen mode with the Nagarkot Forwarders branding.

> **Note:** On the first launch, Windows may show a security warning ("Windows protected your PC"). Click **"More info"** → **"Run anyway"** to proceed. This happens because the EXE is not digitally signed.

---

### 2. The Workflow (Step-by-Step)

#### Step 1 — Select Bill of Entry PDF

1. Click the **"Browse..."** button next to the **"Bill of Entry PDF"** field.
2. A file dialog will open. Navigate to the folder containing your Bill of Entry PDF.
3. Select the PDF file and click **Open**.

**What happens automatically:**
- The app reads the first page of the PDF to detect the **BE Number**.
- The **"Save Output To"** field is automatically populated with a suggested output path:
  `Triple_Duty_Bond_BE_[BE_NUMBER].docx` (saved in the same folder as the input PDF).
- If the BE Number cannot be detected, the output file will be named using the PDF filename instead.

#### Step 2 — Verify Output Location (Optional)

- The output path is pre-filled for you. If you want to save the document to a different folder or with a different name:
  1. Click the **"Browse..."** button next to **"Save Output To"**.
  2. Choose a location and filename.
  3. Click **Save**.

#### Step 3 — Change Template (Optional — Advanced)

The default Word template (`Triple_Duty_Bond_template__IR53112.docx`) is bundled inside the application. You do **not** need to change it under normal circumstances.

If you need to use a different template:
1. Check the **"Change template file"** checkbox.
2. A new row will appear with a **"Browse..."** button.
3. Select the replacement `.docx` template file.

> **Important:** The custom template must contain the same placeholder text as the original template (e.g., `7281285`, `Rs.24,50,000/-`, `1855456`, etc.), otherwise the replacements will not work correctly.

#### Step 4 — Generate the Document

1. Click the **"⚙ Generate Triple Duty Bond"** button.
2. The **Process Logs** panel at the bottom will display real-time progress:
   - **[1] Extracting data from PDF** — Shows all extracted fields.
   - **[2] Calculating Bond Amount** — Shows the debt amount, rounded bond amount, and the amount in words.
   - **[3] Generating document** — Shows each text replacement applied.
3. On success, a confirmation dialog will appear showing a summary of the generated document.
4. Click **"Yes"** to open the output folder directly in Windows Explorer, or **"No"** to dismiss.

---

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Bill of Entry PDF** | The source PDF file from which data is extracted. | `.pdf` file — must be a standard ICEGATE Bill of Entry |
| **Browse... (PDF)** | Opens a file picker to select the input PDF. | Click to open file dialog |
| **Save Output To** | The full path where the generated Word document will be saved. Auto-filled after selecting a PDF. | `.docx` file path |
| **Browse... (Output)** | Opens a Save As dialog to choose a custom output location. | Click to open save dialog |
| **Change template file** | Checkbox — reveals the template file selector when checked. | Check/Uncheck |
| **Word Template** | Path to the Word template used for generation. Pre-filled with the bundled default template. | `.docx` file path |
| **Browse... (Template)** | Opens a file picker to select a custom template. Only visible when the checkbox is checked. | Click to open file dialog |
| **⚙ Generate Triple Duty Bond** | The main action button. Triggers the full extraction → calculation → generation pipeline. | Click to generate |
| **Process Logs** | A read-only text area that displays step-by-step progress, extracted values, and any errors encountered. | Auto-populated during generation |

---

## Understanding the Output

When generation is successful, the Process Logs will display a detailed summary. Here is what each extracted field means:

| Extracted Field | Source in PDF | Used For |
| :--- | :--- | :--- |
| **BE Number** | Found after `INBOM4` on the first page | Inserted into the bond document as the Bill of Entry reference number |
| **BE Date** | Found after the BE Number (format: `DD/MM/YYYY`) | Converted to `DD-Mon-YYYY` format and inserted into all date references |
| **Importer Name** | Found after `1.IMPORTER NAME` | Not directly inserted (retained in logs for verification) |
| **IEC No** | Found after `IEC/Br` | Not directly inserted (retained in logs for verification) |
| **WH Code** | Warehouse code pattern (e.g., `ESA4A001`) | Not directly inserted (retained in logs for verification) |
| **Debt Amount** | Found after `INBOM4 WH` | This is already 3× the total duty. Rounded up to nearest ₹5,000 to get the Bond Amount |
| **TOT.ASS VAL** | 8th value on the line after `8.G.CESS 18.TOT.ASS VAL` | Inserted into the "Value Rs." column of the bond table |
| **Invoice Count** | Found after `TYPE INV ITEM CONT ... Nos` | Inserted into the "Sl No of Invoice" cell of the bond table |
| **Packages** | Found after `BE PKG` | Inserted into the "Packages" cell of the bond table (e.g., `3 PKG`) |

### Bond Amount Calculation

The bond amount is calculated as follows:

```
1. Read DEBT AMT from the PDF (this value is already 3× the total duty)
2. Round UP to the nearest ₹5,000
   Example: DEBT AMT = ₹2,43,567 → Bond Amount = ₹2,45,000
3. Format in Indian numbering: Rs.2,45,000/-
4. Convert to words: "Two Lakh Forty-Five Thousand"
```

### Output File

- **Format:** Microsoft Word Document (`.docx`)
- **Location:** The path shown in the "Save Output To" field
- **Naming Convention:** `Triple_Duty_Bond_BE_[BE_NUMBER].docx`
- **Formatting:** All bold, italic, and font styles from the original template are preserved

---

## Troubleshooting & Validations

If you encounter an error, check this table:

| Message | What It Means | Solution |
| :--- | :--- | :--- |
| `Please select a Bill of Entry PDF file.` | The "Bill of Entry PDF" field is empty. You clicked Generate without selecting a file. | Click **Browse...** next to the PDF field and select a valid PDF file. |
| `Please select a template file.` | The template path is empty or was cleared. | Uncheck and re-check "Change template file", or restart the application to reload the bundled template. |
| `Please specify an output file location.` | The "Save Output To" field is empty. | Click **Browse...** next to the output field and choose a save location, or re-select the PDF file (which auto-fills this). |
| `PDF file not found: [path]` | The file you selected no longer exists at the specified path (it may have been moved, renamed, or deleted). | Click **Browse...** again and re-select the correct PDF file. |
| `Template file not found: [path]` | The template file cannot be found. This can happen if you selected a custom template that was moved. | Uncheck "Change template file" to revert to the bundled default, or browse to the correct template location. |
| `⚠ Template not found – please browse to select` | Shown at launch — the bundled template could not be located. This may happen if the EXE was improperly built. | Use the "Change template file" checkbox and manually browse to a valid `.docx` template. |
| `❌ ERROR: [detailed message]` | An unexpected error occurred during extraction or document generation. The full Python traceback is printed in the Process Logs. | Check the logs for specifics. Common causes: corrupted PDF, non-standard BE format, or the output path is read-only. See below. |

### Common Scenarios

**The extracted values look wrong or are showing as `0` / empty:**
- The Bill of Entry PDF may not follow the standard ICEGATE format that the tool expects.
- The tool specifically looks for patterns like `INBOM4 [BE_NUMBER] [DATE]` and `INBOM4 WH [DEBT_AMT]`. If your PDF uses a different port code or format, the regex patterns will not match.

**The output document looks the same as the template:**
- This means the placeholder text in the template did not match the expected values (`7281285`, `Rs.24,50,000/-`, `1855456`, etc.).
- If you are using a custom template, ensure it contains the exact placeholder strings the application looks for.

**"Permission denied" or unable to save:**
- The output folder may be read-only, or the file may already be open in Microsoft Word.
- Close the previously generated document in Word and try again, or choose a different output folder.

**The application window appears blank or too small:**
- Try maximizing the window manually by pressing `Win + ↑` or double-clicking the title bar.

---

## Quick Reference Card

```
┌──────────────────────────────────────────────────┐
│  TRIPLE DUTY BOND GENERATOR — Quick Steps        │
├──────────────────────────────────────────────────┤
│                                                  │
│  1. Browse → Select Bill of Entry PDF            │
│  2. Verify output path (auto-filled)             │
│  3. Click "⚙ Generate Triple Duty Bond"         │
│  4. Review the Process Logs                      │
│  5. Click "Yes" to open the output folder        │
│  6. Open the .docx and verify before use         │
│                                                  │
└──────────────────────────────────────────────────┘
```

---

*Nagarkot Forwarders Pvt. Ltd. ©*
