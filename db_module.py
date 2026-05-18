"""
Supabase REST API wrapper for the deals table.

Uses stdlib only (urllib) — no extra dependencies. All operations are best-effort:
DB failures never block the contract generator. If env vars are missing, the
module degrades gracefully and all functions return None / empty.

Configured via env vars:
  SUPABASE_URL   — e.g. https://xxxxxxxx.supabase.co
  SUPABASE_KEY   — secret/service_role key (server-only, never sent to browser)
"""
import os
import json
import urllib.request
import urllib.parse
import urllib.error


SUPABASE_URL = (os.environ.get('SUPABASE_URL', '') or '').rstrip('/')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '') or ''
TABLE = 'deals'
_TIMEOUT = 8  # seconds — keep short so server stays responsive even if Supabase is slow


def is_configured():
    return bool(SUPABASE_URL and SUPABASE_KEY)


def _headers(extra=None):
    h = {
        'apikey': SUPABASE_KEY,
        'Authorization': f'Bearer {SUPABASE_KEY}',
        'Content-Type': 'application/json',
    }
    if extra:
        h.update(extra)
    return h


def _request(method, path, body=None, params=None, headers_extra=None):
    """Low-level wrapper. Returns (status, parsed_body) or (None, None) on error."""
    if not is_configured():
        return None, None
    url = f'{SUPABASE_URL}/rest/v1/{path.lstrip("/")}'
    if params:
        url += '?' + urllib.parse.urlencode(params, doseq=True, safe='*.,()')
    data = None
    if body is not None:
        data = json.dumps(body).encode('utf-8')
    req = urllib.request.Request(url, data=data, method=method, headers=_headers(headers_extra))
    try:
        with urllib.request.urlopen(req, timeout=_TIMEOUT) as r:
            raw = r.read()
            try:
                parsed = json.loads(raw) if raw else None
            except Exception:
                parsed = None
            return r.status, parsed
    except urllib.error.HTTPError as e:
        try:
            err_body = e.read().decode('utf-8', errors='replace')
        except Exception:
            err_body = ''
        print(f'[db] HTTP {e.code} on {method} {path}: {err_body[:200]}')
        return e.code, None
    except Exception as e:
        print(f'[db] error on {method} {path}: {e}')
        return None, None


# ────────────────────────────────────────────────────────────────────────────
# Public API
# ────────────────────────────────────────────────────────────────────────────
def save_deal(form_data, entity, deal_type):
    """
    Insert a new deal row. Called after a successful contract generation.
    Best-effort — returns the new row dict or None on failure.

    entity:    'FundGate' or 'Fundkey'
    deal_type: 'weekly' or 'daily'
    """
    if not is_configured():
        return None

    def _money(v):
        try:
            return float(str(v).replace('$', '').replace(',', '').replace('%', ''))
        except Exception:
            return None

    def _date(v):
        if not v:
            return None
        s = str(v).strip()
        # Accept m/d/Y, m/d/y, Y-m-d
        from datetime import datetime
        for fmt in ('%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d'):
            try:
                return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
            except Exception:
                pass
        return None

    row = {
        'merchant_legal_name': (form_data.get('Merchant_Legal_Name') or '').strip(),
        'merchant_dba':        (form_data.get('Merchant_DBA') or '').strip(),
        'entity':              entity,
        'deal_type':           deal_type,
        'state_of_org':        (form_data.get('State_of_Organization') or '').strip().upper(),
        'purchase_price':      _money(form_data.get('Purchase_Price')),
        'purchased_amount':    _money(form_data.get('Purchased_Amount')),
        'agreement_date':      _date(form_data.get('Agreement_Date')),
        'full_data':           form_data,
    }
    # Don't save junk rows with no merchant name
    if not row['merchant_legal_name']:
        return None

    status, body = _request(
        'POST', TABLE, body=row,
        headers_extra={'Prefer': 'return=representation'},
    )
    if status and 200 <= status < 300 and isinstance(body, list) and body:
        return body[0]
    return None


def search_deals(query='', entity=None, limit=50):
    """
    Search by merchant name (case-insensitive partial match). Optionally
    filter by entity. Returns most recent first.
    """
    if not is_configured():
        return []

    params = [
        ('select', 'id,created_at,merchant_legal_name,merchant_dba,entity,deal_type,state_of_org,purchase_price,purchased_amount,agreement_date'),
        ('order', 'created_at.desc'),
        ('limit', str(max(1, min(limit, 200)))),
    ]
    q = (query or '').strip()
    if q:
        # ilike with wildcards — case-insensitive partial match on either field
        params.append(('or', f'(merchant_legal_name.ilike.*{q}*,merchant_dba.ilike.*{q}*)'))
    if entity:
        params.append(('entity', f'eq.{entity}'))

    status, body = _request('GET', TABLE, params=params)
    if status and 200 <= status < 300 and isinstance(body, list):
        return body
    return []


def get_deal(deal_id):
    """Fetch a single deal by id (including full_data for cloning)."""
    if not is_configured() or not deal_id:
        return None
    params = [('id', f'eq.{deal_id}'), ('select', '*'), ('limit', '1')]
    status, body = _request('GET', TABLE, params=params)
    if status and 200 <= status < 300 and isinstance(body, list) and body:
        return body[0]
    return None


def delete_deal(deal_id):
    """Delete a single deal by id."""
    if not is_configured() or not deal_id:
        return False
    params = [('id', f'eq.{deal_id}')]
    status, _ = _request('DELETE', TABLE, params=params)
    return bool(status and 200 <= status < 300)
