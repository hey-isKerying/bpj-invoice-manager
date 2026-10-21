# BPJ Invoice Manager

## Deployment
1. Create a Supabase project. In SQL Editor run `supabase.sql`.
2. Upload the **contents** of this folder to the root of a new GitHub repository. Preserve `templates/` and `api/`.
3. Import repository into Vercel (Framework Preset: **Other**, root directory `./`). Add environment variables:
   - `SUPABASE_URL` = Project URL (Supabase Settings → API)
   - `SUPABASE_SERVICE_ROLE_KEY` = service_role key (Supabase API Keys; **SECRET**, never add to index.html or GitHub)
   - `ADMIN_PASSWORD` = strong unique password (not your Supabase password)
4. Deploy and visit your Vercel domain. Sign in with `ADMIN_PASSWORD`.

## Supported
- Five supplied Excel templates preserved as original workbooks, including embedded garment images.
- Invoice counter starts at 0021, increments transactionally on successful save; previewed next number is indicative.
- Customer/date/items, calculations, Excel downloads, searchable history, monthly charts, monthly ZIP of Excel files + JSON manifest.
- Records are stored in Supabase database. Generated Excel files are **not** stored in Supabase Storage, saving storage capacity.

## IMPORTANT: PDF conversion not yet included
This project **does not yet generate PDF**. High-fidelity Excel-to-PDF conversion requires a separately hosted LibreOffice/Gotenberg service or equivalent; Vercel's Python serverless environment does not include LibreOffice. Do not promise pixel-perfect PDF or ZIP containing PDF until that service is integrated and all five templates are print-tested. You can open the exported Excel and use Excel's Print/Export as PDF in the meantime.

## Limitations / safety
- Single-admin password access; use strong password. A shared password is not a multi-user role system. The browser retains the password in sessionStorage until the tab/session closes.
- Monthly ZIP contains Excel + manifest, **not PDF**. No deletion endpoint is provided: deletion should only be implemented after backup verification and a recovery strategy. Database records are compact.
- Excel formulas are evaluated when Excel opens the file; openpyxl does not calculate cached values. PDF conversion service should recalculate formulas.
- Some templates have a reserved printed size disclaimer in rows 27–28; avoid entering descriptions that overflow into this disclaimer. Validate A4 page setup and artwork positioning with a physical test print.
- Current history listing is capped at 1000 records and monthly ZIP at 1000 records. Pagination/batch export required if you exceed this.
- Only editable fields are filled in existing templates; original source workbook image content is not redrawn. Saving through openpyxl may change unsupported Excel drawing metadata; visually compare each exported template before production use.
