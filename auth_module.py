"""
Lightweight auth for the Past Deals area only.

Design:
  - Single shared password from env var SITE_PASSWORD (default: 'admin')
  - Successful login sets a signed cookie that lasts 7 days
  - Only routes that read historical deal data check the cookie
  - Contract generation routes stay public (no friction for daily use)

Cookie format: HMAC-signed value of the form "<expiry_unix_ts>.<sha256_hex>"
Verified server-side using a secret derived from SITE_PASSWORD.
"""
import os
import time
import hmac
import hashlib


PASSWORD = os.environ.get('SITE_PASSWORD', 'admin')
COOKIE_NAME = 'fg_session'
COOKIE_DAYS = 7
_SECRET = hashlib.sha256(('fundgate-deals-v1::' + PASSWORD).encode()).digest()


def _sign(expiry_ts):
    msg = str(expiry_ts).encode()
    sig = hmac.new(_SECRET, msg, hashlib.sha256).hexdigest()
    return f'{expiry_ts}.{sig}'


def make_cookie_value():
    expiry = int(time.time()) + COOKIE_DAYS * 86400
    return _sign(expiry)


def cookie_header():
    """Build the Set-Cookie header value for a fresh login."""
    val = make_cookie_value()
    return (
        f'{COOKIE_NAME}={val}; '
        f'Max-Age={COOKIE_DAYS * 86400}; '
        f'Path=/; '
        f'HttpOnly; '
        f'SameSite=Lax'
    )


def clear_cookie_header():
    return f'{COOKIE_NAME}=; Max-Age=0; Path=/; HttpOnly; SameSite=Lax'


def is_valid_cookie(cookie_header_str):
    """Given a raw Cookie: header string, return True if a valid, unexpired session cookie is present."""
    if not cookie_header_str:
        return False
    for part in cookie_header_str.split(';'):
        kv = part.strip().split('=', 1)
        if len(kv) != 2:
            continue
        name, val = kv[0].strip(), kv[1].strip()
        if name != COOKIE_NAME:
            continue
        if '.' not in val:
            continue
        try:
            expiry_str, sig = val.split('.', 1)
            expiry_ts = int(expiry_str)
        except Exception:
            continue
        # Expired?
        if expiry_ts < int(time.time()):
            continue
        # Signature valid?
        expected = hmac.new(_SECRET, str(expiry_ts).encode(), hashlib.sha256).hexdigest()
        if hmac.compare_digest(expected, sig):
            return True
    return False


def check_password(submitted):
    """Constant-time compare of a submitted password against the configured one."""
    if submitted is None:
        return False
    return hmac.compare_digest(str(submitted).encode(), PASSWORD.encode())
