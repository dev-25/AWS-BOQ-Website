# AWS Pricing BOQ Generator

Link - https://aws-boq-preparation.streamlit.app

A local Streamlit app that turns an AWS Pricing Calculator estimate into a
customer-ready Excel BOQ, formatted like the Dixit Infotech pricing template
(Estimate summary, Summary/commercials in INR, Detailed Estimate, Note, Terms
and Conditions).

## What it does

1. **Pulls the AWS line items** two ways:
   - **Fetch from link** — paste the `calculator.aws/#/estimate?id=...` share
     link and click *Fetch estimate data*. This calls the same public,
     unauthenticated endpoint the calculator.aws website itself uses to load
     a saved estimate. It is **not an official/documented AWS API**, so if
     AWS changes it, the fetch may stop working — that's what tab 2 is for.
   - **Upload exported CSV** — on the calculator.aws "My estimate" page, use
     **Export → CSV**, then upload that file. This is the official AWS export
     and will always work.
   - Either way, the line items land in an editable table — add, remove, or
     fix rows by hand before generating the file.
2. **Manual commercial inputs**: Dollar rate (USD→INR), MSP as a **percentage**
   of the AWS Total Estimate Monthly cost (e.g. enter `10` for 10%, and Excel
   computes the ₹ amount for you), one-time Deployment & Migration, FW
   License (Annual), and EDR License (Annual) — each entered by hand.
3. **Note** and **Terms and Conditions** are optional, editable blocks. Their
   position in the sheet is computed automatically from however many detail
   rows you end up with — they're never pinned to a fixed row number, so a
   5-row estimate and a 50-row estimate both lay out correctly.
4. **Download** — generates a formatted `.xlsx` with the same fonts, colors,
   borders, number formats (₹ / $) and live formulas as the original
   template (totals are calculated by Excel formulas, not hardcoded, so they
   stay correct if you tweak a number after downloading). Default Excel
   gridlines are turned off so the sheet is clean white, with visible
   borders only around the actual content tables. Column F (Upfront) is
   hidden by default since it's almost always $0. The Detailed Estimate
   table is bounded to column J, and there are two blank rows of breathing
   room before the Detailed Estimate, Note, and Terms sections.

## Starting a new proposal

Click **\U0001F504 New** at the top right at any time to wipe the current
estimate (line items, link, all manual inputs) and start completely fresh —
no leftover data from the previous proposal.

## Setup (Mac mini M2)

```bash
cd aws_boq_app
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

This opens the app at `http://localhost:8501` in your browser. Streamlit
runs entirely on your Mac — nothing here needs an AWS account or
credentials.

## Notes

- The "Fetch from link" endpoint is the same one used by a public AWS
  sample MCP server (`aws-samples/sample-aws-pricing-calculator-mcp`); it's
  unauthenticated and undocumented, so treat it as best-effort. The CSV
  upload path is the guaranteed-accurate fallback.
- Upfront/Monthly are pulled per line item; "First 12 months total" is
  always computed in Excel as `Upfront + Monthly × 12`.
- The summary's "AWS Total Estimate" (in USD and INR) is also computed by
  formula as the sum of all detail rows, so it always matches the table —
  even after you add/edit rows in the app or later in Excel.
- Re-run `streamlit run app.py` any time you want to build a new estimate;
  each session starts fresh.
