
# Test Fixture Management (Flask + Excel)

## Quick start (VS Code)
1. Open this folder in VS Code: `/mnt/data/test_fixture_manager`
2. Create/activate a virtual environment
   - Windows: `py -m venv .venv && .venv\Scripts\activate`
   - macOS/Linux: `python3 -m venv .venv && source .venv/bin/activate`
3. Install dependencies: `pip install -r requirements.txt`
4. Run the backend: `python app.py`
5. Open the UI at: http://localhost:5000

## Data files
- Source inventory: `data/Test Fixture Location_final.xlsx`
- Borrow log (auto-created): `data/borrowed_test_fixtures.xlsx`

## Flow
- Search by Article → choose system (SAFT, VSFT, VSICT, SPEA3030) → see details and available units
- Borrow: enter name/id (batch), click Check Out → log is appended and availability decreases
- Return: simulate QR by using Check In; backend will return all outstanding borrows for that article/system

> Availability = Excel "Available Units (Qty.)" minus outstanding borrowed quantities.

## Notes
- "SPEA3030" is detected when fixture description contains "SPEA".
- You can customize mapping in `system_label()` in `app.py` if needed.
