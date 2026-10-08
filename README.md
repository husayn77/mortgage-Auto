# MGT Mapper

A Next.js app that takes an input Excel file of member details and generates the
**MGT formula workbook** (`EXCEL TEMPLATE.xlsx`), filling the `RUN` sheet with
live Excel formulas and amount-in-words columns.

## Run it

```powershell
npm install
npm run dev      # http://localhost:3000
# or for production:
npm run build
npm run start
```

Open the page, drag in your input Excel, and click **Generate MGT template** to
download the filled workbook.

## How it works

- The template lives at `templates/EXCEL TEMPLATE.xlsx`.
- Column headers in your input are auto-detected (case-insensitive). Recognised
  fields: Bank/SN, Account Number, Account Name, PIN/Pension No, Balance,
  Address, Next of Kin, NIN, BVN, Phone, DOB, Employer, Signatory.
- Per row, the app writes the raw values and these **live formulas** into `RUN`:

  | Col | Meaning            | Formula            |
  |-----|--------------------|--------------------|
  | G   | Purchase price     | `=F/4+200000`      |
  | I   | Equity Contribution| `=F/4`             |
  | K   | Forsale value      | `=G-300000`        |
  | M   | Sum Insured        | `=G`               |
  | N   | Fire Insurance     | `=M*0.21%`         |
  | O   | Credit Insurance   | `=200000*1%`       |
  | P   | Processing fee     | `=I*2%`            |
  | W   | Year of birth      | `=YEAR(V)`         |
  | X   | Age                | `=<currentYear>-W` |
  | Z   | Residential address| `=Q`               |

- Columns `H`, `J`, `L` are the Naira/Kobo amount-in-words for `G`, `I`, `K`
  (computed in code, since Excel has no built-in number-to-words).

## Adjusting the mapping

Edit `src/lib/mapping.ts`:

- `HEADER_ALIASES` — add alternative header names for a field.
- `RAW_COLUMN` — which template column each raw field is written to.
- `FORMULA_COLUMN` — the formulas per column.

A sample input is provided at `_test/sample-input.xlsx` (regenerate with
`node _test/gen.cjs`).

## Welcome letters (Word)

The **Welcome Letters** tab fills a Word mail-merge template with each member's
details and downloads a ZIP of one `.docx` per member, named
`<member_name>_<template>.docx` (e.g. `ahmed_sani_premium.docx`).

Choose the template (Premium Pension / Fidelity Pension) from the dropdown
before generating. Field names are matched case/spelling-insensitively, so
minor differences between templates still fill correctly. Embedded images are
recompressed automatically to keep file sizes small.

### Adding another letter template

1. Drop your `.docx` into the `templates/` folder.
2. (Optional) add an entry to `templates/letter-templates.json` for a friendly
   label and filename suffix:

   ```json
   [
     { "file": "MY TEMPP.docx", "label": "Premium Pension", "suffix": "premium" },
     { "file": "MY TEMPP FID.docx", "label": "Fidelity Pension", "suffix": "fidelity" },
     { "file": "My New Template.docx", "label": "ARM Pension", "suffix": "arm" }
   ]
   ```

   If you skip the JSON entry, the template still appears, using its filename as
   the label. The template should use Word mail-merge fields (`«FieldName»`)
   such as `ACCOUNT_NAME`, `PIN_NUMBER`, `Purchase_price`, `Equity_Contribution`,
   `Forsale_value`, `Processing_fee_`, etc.
