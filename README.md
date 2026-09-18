# Restopros Recon Tracker

Turns an **Xactimate Internal TAM PDF** into a Restopros cost / royalty / profitability workbook (same model as the 600 Elgin recon sheet).

## Required upload

The uploaded file **must be an Xactimate PDF Internal TAM report** with **all 7 print selections checked**:

1. Coversheet
2. Line item detail
3. Summary
4. Recap by room
5. Recap by category
6. Labor breakdown
7. Sketch / Grand total areas

In Xactimate: **Documents → Reports → Report type = Internal TAM → Print selections → check all → export PDF**.

Final Draft, Scope, Abbreviated, or a TAM printed without Labor breakdown will be rejected.

## Run

```bash
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5055

## Defaults

| Assumption | Default |
|---|---|
| Supervisor gross | $30/hr (CLN-S / supervisory) |
| Worker gross | $22/hr (all other TAM labor codes) |
| Burden | 18% |
| Royalty | 8% of negotiated price |
| Contingency | 5% of materials + equipment |

Negotiated price defaults to Replacement Cost Value and is the yellow box on the Summary tab.
