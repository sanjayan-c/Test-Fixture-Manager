from flask import Flask, jsonify, request, send_from_directory
from pathlib import Path
from datetime import datetime
import pandas as pd
import uuid

# Serve files from the folder that contains this script
BASE = Path(__file__).parent.resolve()
app = Flask(__name__, static_folder=str(BASE), static_url_path='')

DATA_DIR = BASE / 'data'
FIX_FILE = DATA_DIR / 'Test Fixture Location_final.xlsx'
BORROW_FILE = DATA_DIR / 'borrowed_test_fixtures.xlsx'

# ---------- Data helpers ----------
def load_fixtures():
    """
    Excel:
      Sheet name: 'Fixtures'
      First row is the real header row.
      Key columns: Article, Part Number, Name, Fixture Type, Fixture Description, Location, Available Units (Qty.)
    """
    xls = pd.ExcelFile(FIX_FILE)
    df = xls.parse('Fixtures')
    header = df.iloc[0].tolist()
    df = df.iloc[1:].copy()
    df.columns = header

    # Normalize
    df['Article'] = df['Article'].astype(str).str.strip()
    df['Fixture Type'] = df['Fixture Type'].astype(str).str.strip()
    if 'Fixture Description' in df.columns:
        df['Fixture Description'] = df['Fixture Description'].astype(str)
    if 'Available Units (Qty.)' in df.columns:
        df['Available Units (Qty.)'] = pd.to_numeric(
            df['Available Units (Qty.)'], errors='coerce'
        ).fillna(0).astype(int)
    else:
        df['Available Units (Qty.)'] = 0
    return df

def system_label(row):
    ft = (row.get('Fixture Type') or '').upper()
    desc = (row.get('Fixture Description') or '').upper()
    if 'VSFT' in ft: return 'VSFT'
    if 'VSICT' in ft: return 'VSICT'
    if 'SAFT' in ft: return 'SAFT'
    if 'SPEA' in desc: return 'SPEA3030'
    return ft or 'OTHER'

def ensure_borrow_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Target schema (System retained internally to keep per-system availability exact):
      borrow_id, Article, Part Number, System, Quantity,
      Client Name, Client Phone, Location,
      Borrowed At, Returned At
    """
    columns = [
        'borrow_id','Article','Part Number','System','Quantity',
        'Client Name','Client Phone','Location',
        'Borrowed At','Returned At'
    ]
    for c in columns:
        if c not in df.columns:
            df[c] = pd.NA
    # Drop legacy columns we don't use anymore (safe no-op if missing)
    legacy_to_drop = ['Name','Employee Name','Employee Number']
    for c in legacy_to_drop:
        if c in df.columns:
            df = df.drop(columns=[c])
    # Reorder
    df = df[columns]
    return df

def load_borrow():
    if not BORROW_FILE.exists():
        cols = [
            'borrow_id','Article','Part Number','System','Quantity',
            'Client Name','Client Phone','Location',
            'Borrowed At','Returned At'
        ]
        pd.DataFrame(columns=cols).to_excel(BORROW_FILE, index=False)
    df = pd.read_excel(BORROW_FILE)
    return ensure_borrow_schema(df)

def save_borrow(df: pd.DataFrame):
    ensure_borrow_schema(df).to_excel(BORROW_FILE, index=False)

def availability(df: pd.DataFrame, article: str, system: str) -> int:
    """Available = Excel qty minus currently-open borrows for that article+system."""
    sys_norm = (system or '').upper()
    base = df[(df['Article'] == str(article)) & (df.apply(system_label, axis=1).str.upper() == sys_norm)]
    base_qty = int(base['Available Units (Qty.)'].sum()) if not base.empty else 0

    bor = load_borrow()
    if bor.empty:
        return base_qty

    mask_open = (
        (bor['Article'].astype(str) == str(article)) &
        (bor['System'].astype(str).str.upper() == sys_norm) &
        (bor['Returned At'].isna())
    )
    used = int(pd.to_numeric(bor.loc[mask_open, 'Quantity'], errors='coerce').fillna(0).sum()) if mask_open.any() else 0
    return max(base_qty - used, 0)

# ---------- Routes ----------
@app.route('/')
def index():
    return send_from_directory(str(BASE), 'index.html')

@app.get('/api/search')
def api_search():
    """
    ?article=...
    Finds matching article (exact first, then contains).
    Returns ONLY systems that exist for that article AND have availability > 0.
    If multiple distinct articles match on 'contains', returns a choice list.
    """
    article = request.args.get('article', '').strip()
    if not article:
        return jsonify(found=False, error="Missing article"), 400

    df = load_fixtures()

    # exact first
    sub = df[df['Article'].astype(str).str.fullmatch(str(article))]
    if sub.empty:
        # fallback: contains
        sub = df[df['Article'].astype(str).str.contains(str(article), na=False)]

        uniq_articles = sub['Article'].dropna().astype(str).unique().tolist()
        if len(uniq_articles) > 1:
            choices = (sub[['Article', 'Part Number', 'Name']]
                       .astype(str)
                       .drop_duplicates()
                       .head(20)
                       .to_dict(orient='records'))
            return jsonify(found='multiple', choices=choices)

    if sub.empty:
        return jsonify(found=False)

    row0 = sub.iloc[0]
    chosen_article = str(row0.get('Article', ''))

    # Dynamic systems present (no hardcoded list)
    systems_present = sorted(set(sub.apply(system_label, axis=1).tolist()))

    systems_payload = []
    for s in systems_present:
        avail = availability(df, chosen_article, s)
        if avail > 0:
            systems_payload.append({'system': s, 'available_units': avail})

    return jsonify(
        found=True,
        article=chosen_article,
        part_number=str(row0.get('Part Number','')),
        name=str(row0.get('Name','')),
        systems=systems_payload
    )

@app.get('/api/details')
def api_details():
    """
    ?article=...&system=...
    Returns consolidated details + live availability, including locations list.
    """
    article = request.args.get('article', '').strip()
    system = request.args.get('system', '').strip()
    if not article or not system:
        return jsonify(error='Missing params'), 400

    sys_norm = system.upper()
    df = load_fixtures()
    sub = df[(df['Article'].astype(str) == article) &
             (df.apply(system_label, axis=1).str.upper() == sys_norm)]
    if sub.empty:
        return jsonify(error='Not found'), 404

    avail = availability(df, article, sys_norm)
    locations = sub['Location'].dropna().astype(str).unique().tolist() if 'Location' in sub.columns else []
    row = sub.iloc[0]

    return jsonify({
        'article': article,
        'part_number': str(row.get('Part Number','')),
        'name': str(row.get('Name','')),
        'system': sys_norm,
        'available_units_total': int(avail),
        'locations': locations,
        'primary_location': locations[0] if locations else '',
        'description': str(row.get('Fixture Description',''))
    })

@app.post('/api/borrow')
def api_borrow():
    """
    Body: { article, system, quantity, client_name, client_phone, location }
    Writes a row to borrowed_test_fixtures.xlsx and enforces availability.
    """
    data = request.get_json(force=True, silent=True) or {}
    article = str(data.get('article','')).strip()
    system  = str(data.get('system','')).strip()
    loc     = str(data.get('location','')).strip()
    qty_raw = data.get('quantity', 1)
    try:
        qty = int(qty_raw)
    except Exception:
        qty = 0
    client_name  = str(data.get('client_name','')).strip()
    client_phone = str(data.get('client_phone','')).strip()

    if not (article and system and qty > 0 and client_name and client_phone):
        return jsonify(ok=False, error='Missing required fields'), 400

    df = load_fixtures()
    if qty > availability(df, article, system):
        return jsonify(ok=False, error='Not enough units available'), 400

    # enrich part number
    sub = df[df['Article'].astype(str) == article]
    part = str(sub.iloc[0].get('Part Number','')) if not sub.empty else ''

    # Default location to primary for that system if none provided
    if not loc:
        sys_norm = system_label(sub.iloc[0]) if not sub.empty else system.upper()
        sub_sys = sub[sub.apply(system_label, axis=1).str.upper() == sys_norm]
        loc = (sub_sys['Location'].dropna().astype(str).iloc[0]
               if ('Location' in sub_sys.columns and not sub_sys.empty and not sub_sys['Location'].dropna().empty)
               else '')

    bor = load_borrow()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    bid = str(uuid.uuid4())
    new = {
        'borrow_id': bid,
        'Article': article,
        'Part Number': part,
        # Keep System internally to keep per-system availability correct
        'System': system.upper(),
        'Quantity': qty,
        'Client Name': client_name,
        'Client Phone': client_phone,
        'Location': loc,
        'Borrowed At': now,
        'Returned At': pd.NaT
    }
    bor = pd.concat([bor, pd.DataFrame([new])], ignore_index=True)
    save_borrow(bor)
    return jsonify(ok=True, borrow_id=bid, article=article, part_number=part,
                   system=system.upper(), quantity=qty, location=loc, timestamp=now)

@app.post('/api/return')
def api_return():
    """
    Body: { borrow_ids: [ ... ] }  or  { borrow_id: "..." }
    Marks those rows as returned (sets Returned At = now) if they are currently open.
    """
    data = request.get_json(force=True, silent=True) or {}
    ids = data.get('borrow_ids')
    if not ids:
        one = data.get('borrow_id')
        ids = [one] if one else []
    ids = [str(x).strip() for x in ids if str(x).strip()]

    if not ids:
        return jsonify(ok=False, error='Missing borrow_id(s)'), 400

    bor = load_borrow()
    if bor.empty:
        return jsonify(ok=False, error='No borrow records'), 404

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    mask = bor['borrow_id'].astype(str).isin(ids) & (bor['Returned At'].isna())
    updated_count = int(mask.sum())
    if updated_count == 0:
        return jsonify(ok=False, error='No open borrows matched those IDs'), 404

    bor.loc[mask, 'Returned At'] = now
    save_borrow(bor)
    return jsonify(ok=True, returned=updated_count, timestamp=now, borrow_ids=ids)

# Fallback to serve static files (index.html in same folder)
@app.route('/<path:path>')
def static_forward(path):
    return send_from_directory(str(BASE), path)

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
