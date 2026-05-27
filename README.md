# Beauty Brand BP Generator

A Streamlit app that turns a Sales Report and Stock Levels Report into:

- `[Brand] BP Data.xlsx`
- `[Brand] Business Analysis.docx`

The app auto-detects SKU, product name, quantity, available stock, incoming stock, and optional catalogue columns. It then builds SABC sales classification, inventory coverage, replenishment actions, Excel summaries, charts, and a Word business analysis report.

## Run

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Login

The app starts with a default admin account:

```text
Username: admin
Password: change-me-now
```

Change this password after the first login. Admin users can add users, delete users, reset user passwords, and change their own password. Passwords are stored as hashes, so the app does not show plaintext passwords.

## Inputs

Required:

- Sales Report Excel file
- Stock Levels Excel file

Optional:

- Product Catalogue / Price List Excel file

## Key Logic

- Sales are aggregated by Product SKU.
- If a Month/Year column exists, the app uses the latest available 12-month window in the file.
- Future Inventory = Available + Incoming.
- Adjusted Future Inventory = MAX(0, Future Inventory).
- Avg Monthly Sales = Qty / 12.
- Coverage = Adjusted Future Inventory / Avg Monthly Sales.
- Inventory status and action are assigned automatically.
