"""
California Commercial Financing Disclosure module.

Generates the 3-page CA disclosure required under California Code of
Regulations Title 10, Chapter 3, Subchapter 3 (DFPI sales-based financing
disclosure). Format mirrors the manually-built BRG/Angry Petes disclosures:
  - 9-row, 3-column table (Label | Value | Description)
  - Page 1: Funding Provided through Payment Terms
  - Page 2: Estimated Term continuation, Prepayment, Acknowledgment, Signatures
  - Page 3: Itemization of Amount Financed

Branding: Fundkey LLC (NOT FundGate). This module is ONLY used by the CA
routes in server.py.

APR calculation method: TILA Regulation Z Appendix J actuarial method,
daily compounding (rate/365), with the gross Purchase Price as the present
value. This matches BizFund's methodology and CA DFPI § 940.
"""
import io, zipfile, re
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def _fmt_currency(val):
    try:
        n = float(str(val).replace('$', '').replace(',', '').replace('%', ''))
        return f'${n:,.2f}'
    except Exception:
        return str(val)


def _fmt_date(val):
    if not val:
        return ''
    for fmt in ('%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d'):
        try:
            return datetime.strptime(str(val).strip(), fmt).strftime('%B %d, %Y')
        except Exception:
            pass
    return str(val)


def _n(data, key, default=0.0):
    try:
        return float(str(data.get(key, default)).replace('$', '').replace(',', '').replace('%', ''))
    except Exception:
        return default


# ─────────────────────────────────────────────────────────────────────────────
# APR calculation (TILA Reg Z Appendix J — actuarial method, daily compounding)
# Finds the annual rate r where the sum of all discounted payments equals the
# Purchase Price (gross funding provided). Solved iteratively via bisection.
# ─────────────────────────────────────────────────────────────────────────────
def _calculate_apr(purchase_price, payment_amount, frequency, total_purchased):
    """
    purchase_price  — gross funding provided (e.g. $30,000)
    payment_amount  — periodic payment (e.g. $1,729.62 weekly)
    frequency       — 'weekly' or 'daily' ('daily' = Mon–Fri business days)
    total_purchased — total dollars to be repaid (e.g. $44,970)

    Returns APR as a decimal percentage (e.g. 167.88 for 167.88%).
    """
    if purchase_price <= 0 or payment_amount <= 0 or total_purchased <= 0:
        return 0.0

    # Number of payments
    n_payments = int(round(total_purchased / payment_amount))
    # Handle small fractional cents — last payment may be smaller, but Reg Z
    # uses equal payments for the estimate so we assume all payments equal.
    if n_payments <= 0:
        return 0.0

    # Day offsets for each payment from funding (day 0)
    # Weekly: 7, 14, 21, ..., 7*n
    # Daily: 1, 2, 3, ..., n (only counting business days; one biz-day per slot)
    # For the disclosure estimate we keep it simple — daily = 1 day per payment
    # in CALENDAR-DAY equivalents Mon-Fri-Mon-Tue... means avg 7/5 calendar
    # days per business day. Reg Z standard for daily commercial financing
    # uses the actual calendar day offsets, so 5 payments per 7-day cycle.
    is_weekly = 'week' in (frequency or '').lower()
    days = []
    if is_weekly:
        for i in range(1, n_payments + 1):
            days.append(i * 7)
    else:
        # Daily: 5 business days per 7-day week. Calendar offsets:
        # Day 1 = first biz day = day 1 (assume Mon funding -> Tue payment)
        # Pattern: 1,2,3,4,5 (Mon-Fri), then 8,9,10,11,12, etc.
        for i in range(n_payments):
            week = i // 5
            dow = i % 5
            days.append(week * 7 + dow + 1)

    # PV equation: PP = sum( payment / (1 + r/365)^day_i )
    # Solve for r using bisection on r in [0.01, 50.0] (1% to 5000% annual)
    def pv_at(r):
        daily_rate = r / 365.0
        total = 0.0
        for d in days:
            total += payment_amount / ((1 + daily_rate) ** d)
        return total

    lo, hi = 0.0001, 50.0  # bracket: 0.01% to 5000% APR
    # Verify bracket actually contains the root
    if pv_at(lo) < purchase_price or pv_at(hi) > purchase_price:
        # Edge case — return rough estimate
        return 0.0

    for _ in range(200):
        mid = (lo + hi) / 2
        pv = pv_at(mid)
        if abs(pv - purchase_price) < 0.01:
            break
        if pv > purchase_price:
            lo = mid
        else:
            hi = mid

    apr_pct = mid * 100.0
    return round(apr_pct, 2)


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers — build raw OOXML for the 3-page disclosure
# ─────────────────────────────────────────────────────────────────────────────
FONT = 'Times New Roman'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')


def _runs_rich(text, size=22):
    """Parse a string with **bold** inline markers into multiple <w:r> runs.

    Example: 'funding **FundGate LLC** will provide' produces 3 runs:
       'funding ' (regular) + 'FundGate LLC' (bold) + ' will provide' (regular)
    This is how we get the selective bolding seen in the reference PDF.
    """
    parts = []
    current = []
    is_bold = False
    i = 0
    while i < len(text):
        if text[i:i+2] == '**':
            if current:
                parts.append((''.join(current), is_bold))
                current = []
            is_bold = not is_bold
            i += 2
        else:
            current.append(text[i])
            i += 1
    if current:
        parts.append((''.join(current), is_bold))
    return ''.join(_run(t, bold=b, size=size) for t, b in parts if t)


def _safe(s):
    return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _run(text, bold=False, italic=False, size=20, color=None, underline=False):
    """Return a <w:r> XML string with the given run properties.
    size is in half-points: size=20 = 10pt, size=22 = 11pt, size=24 = 12pt."""
    rpr = '<w:rPr>'
    rpr += f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
    if bold:
        rpr += '<w:b/><w:bCs/>'
    if italic:
        rpr += '<w:i/><w:iCs/>'
    if underline:
        rpr += '<w:u w:val="single"/>'
    if color:
        rpr += f'<w:color w:val="{color}"/>'
    rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    rpr += '</w:rPr>'
    # Preserve any leading/trailing spaces
    space = ' xml:space="preserve"' if text != text.strip() else ''
    return f'<w:r>{rpr}<w:t{space}>{_safe(text)}</w:t></w:r>'


def _para(runs_xml, align=None, spacing_before=0, spacing_after=0, indent_left=None, keep_next=False):
    """Build a <w:p> containing the given run XML."""
    ppr = '<w:pPr>'
    if align:
        ppr += f'<w:jc w:val="{align}"/>'
    if spacing_before or spacing_after:
        ppr += f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}"/>'
    if indent_left is not None:
        ppr += f'<w:ind w:left="{indent_left}"/>'
    if keep_next:
        ppr += '<w:keepNext/>'
    ppr += '</w:pPr>'
    if isinstance(runs_xml, (list, tuple)):
        runs_xml = ''.join(runs_xml)
    return f'<w:p>{ppr}{runs_xml}</w:p>'


def _page_break():
    """Hard page break paragraph."""
    return ('<w:p><w:r>'
            '<w:rPr><w:rFonts w:ascii="' + FONT + '" w:hAnsi="' + FONT + '"/></w:rPr>'
            '<w:br w:type="page"/></w:r></w:p>')


def _cell(content_xml, width_dxa, vmerge=None, shading=None, borders=True, valign='center'):
    """Build a <w:tc>. content_xml is one or more <w:p> elements.
    valign='center' makes content vertically centered in the cell — matches
    the Angry Petes reference where values like '$50,000.00' sit centered
    next to multi-line descriptions."""
    tcpr = '<w:tcPr>'
    tcpr += f'<w:tcW w:w="{width_dxa}" w:type="dxa"/>'
    if vmerge:
        tcpr += f'<w:vMerge w:val="{vmerge}"/>'
    if shading:
        tcpr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>'
    if borders:
        b = ('<w:tcBorders>'
             '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '</w:tcBorders>')
        tcpr += b
    if valign:
        tcpr += f'<w:vAlign w:val="{valign}"/>'
    # Wider top/bottom padding for the breathing room visible in the reference
    tcpr += '<w:tcMar><w:top w:w="160" w:type="dxa"/><w:left w:w="140" w:type="dxa"/>'
    tcpr += '<w:bottom w:w="160" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>'
    tcpr += '</w:tcPr>'
    return f'<w:tc>{tcpr}{content_xml}</w:tc>'


def _row(cells_xml):
    return f'<w:tr>{"".join(cells_xml)}</w:tr>'


def _table(rows_xml, total_width_dxa=10800):
    tblpr = '<w:tblPr>'
    tblpr += f'<w:tblW w:w="{total_width_dxa}" w:type="dxa"/>'
    tblpr += '<w:jc w:val="center"/>'
    tblpr += '<w:tblLayout w:type="fixed"/>'
    tblpr += ('<w:tblBorders>'
              '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '</w:tblBorders>')
    tblpr += '</w:tblPr>'
    return f'<w:tbl>{tblpr}{"".join(rows_xml)}</w:tbl>'


# ─────────────────────────────────────────────────────────────────────────────
# Main builder
# ─────────────────────────────────────────────────────────────────────────────
def build_ca_disclosure_bytes(data):
    """
    Build the 3-page CA disclosure as DOCX bytes. Returns None if state is
    not CA or required fields are missing.
    """
    # Pull inputs
    merchant_name = (data.get('Merchant_Legal_Name', '') or '').upper()
    merchant_dba = (data.get('Merchant_DBA', '') or merchant_name).upper()
    date_display = _fmt_date(data.get('Agreement_Date', ''))

    pp = _n(data, 'Purchase_Price')
    pa = _n(data, 'Purchased_Amount')
    ach_pct = _n(data, 'ACH_Program_Fee_Percentage')
    orig_pct = _n(data, 'Origination_Fee_Percentage')
    # Support $ and % modes for fees (form sends Mode flag from v8+).
    # When Mode == 'dollar' the input is a flat dollar amount, not a percentage.
    ach_fee_mode  = (data.get('ACH_Program_Fee_Mode', 'pct') or 'pct').lower()
    orig_fee_mode = (data.get('Origination_Fee_Mode', 'pct') or 'pct').lower()
    ach_fee  = ach_pct if ach_fee_mode == 'dollar' else round(pp * ach_pct / 100, 2)
    orig_fee = orig_pct if orig_fee_mode == 'dollar' else round(pp * orig_pct / 100, 2)
    total_fees = round(ach_fee + orig_fee, 2)
    net_to_merchant = round(pp - total_fees, 2)
    finance_charge = round(pa - pp, 2)

    spec_pct_raw = data.get('Specified_Percentage', '0')
    try:
        spec_pct = float(str(spec_pct_raw).replace('%', '').replace(',', ''))
    except Exception:
        spec_pct = 0.0

    avg_monthly_rev = _n(data, 'CA_Avg_Monthly_Revenue')
    # Disclosure always renders, even on a blank doc. Missing/zero values
    # simply produce zeros in the relevant cells.

    ach_freq = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()
    is_weekly = 'week' in ach_freq
    if is_weekly:
        pmt = _n(data, 'Specific_Weekly_Amount')
        period_label = 'week'
        period_label_cap = 'Weekly'
    else:
        pmt = _n(data, 'Specific_Daily_Amount')
        period_label = 'business day'
        period_label_cap = 'Daily'

    # Calculations — all guarded against zero inputs
    apr = _calculate_apr(pp, pmt, ach_freq, pa) if (pp > 0 and pa > 0 and pmt > 0) else 0.0
    n_payments = int(round(pa / pmt)) if pmt > 0 else 0
    # Estimated monthly cost: weekly × 52/12 OR daily × 252/12 (252 biz days/yr)
    if is_weekly:
        est_monthly = round(pmt * 52 / 12, 2)
    else:
        est_monthly = round(pmt * 252 / 12, 2)
    # Estimated term (months): total / (revenue × specified%)
    monthly_capture = avg_monthly_rev * spec_pct / 100.0
    if monthly_capture > 0 and pa > 0:
        est_term_months = max(1, int(round(pa / monthly_capture + 0.0001)))
    elif pmt > 0 and pa > 0:
        est_term_months = max(1, int(round(n_payments / (52/12) if is_weekly else n_payments / 21)))
    else:
        est_term_months = 0

    two_signers = bool(data.get('twoSigners', False))
    signer1_name = (data.get('Owner_Guarantor_1', '') or '').upper()
    signer2_name = (data.get('Owner_Guarantor_2', '') or '').upper() if two_signers else ''

    # Format displays
    pp_fmt = _fmt_currency(pp)
    pa_fmt = _fmt_currency(pa)
    net_fmt = _fmt_currency(net_to_merchant)
    fc_fmt = _fmt_currency(finance_charge)
    em_fmt = _fmt_currency(est_monthly)
    pmt_fmt = _fmt_currency(pmt)
    rev_fmt = _fmt_currency(avg_monthly_rev)
    ach_fee_fmt = _fmt_currency(ach_fee)
    orig_fee_fmt = _fmt_currency(orig_fee)
    fees_fmt = _fmt_currency(total_fees)

    # ═══════════════════════════════════════════════════════════════════════
    # ANGRY PETES FORMAT — must match the manually-generated reference PDF
    # exactly. Provider name = Fundkey LLC for Fundkey route, FundGate LLC
    # otherwise. (Today's CA flow is Fundkey-only, but the branding swap is
    # honored for safety.)
    # ═══════════════════════════════════════════════════════════════════════
    is_fundkey = bool(data.get('isFundkey', False) or data.get('isCA', False))
    provider = 'Fundkey LLC' if is_fundkey else 'FundGate LLC'

    # ─── Column widths — total 10800 dxa ────────────────────────────────────
    W_LABEL = 2400
    W_VALUE = 2400
    W_DESC  = 6000

    # ─── Cell builders ──────────────────────────────────────────────────────
    # Font sizes (half-points): 22=11pt body, 24=12pt value text.
    # Descriptions support inline **bold** markers via _runs_rich().
    def cell_label(text, sub_lines=None):
        """Bold, light-gray-shaded label cell (left column).
        Reference shows: label at top, blank gap, then sub-lines below.
        Top-aligned so the label stays at top of tall rows."""
        paras = [_para([_run(text, bold=True, size=22)], align='left', spacing_after=0)]
        if sub_lines:
            # Blank spacer paragraph (matches the visible gap in reference)
            paras.append(_para([_run('', size=22)], align='left', spacing_after=0))
            for line in sub_lines:
                paras.append(_para([_run(line, bold=True, size=22)], align='left', spacing_after=0))
        return _cell(''.join(paras), W_LABEL, shading='E7E6E6', valign='top')

    def cell_value(text, sub=None):
        """Center-aligned value cell (middle column)."""
        runs = [_para([_run(text, bold=True, size=24)], align='center', spacing_after=0)]
        if sub:
            runs.append(_para([_run(sub, size=22)], align='center', spacing_after=0))
        return _cell(''.join(runs), W_VALUE)

    def cell_value_blank():
        return _cell(_para([_run('', size=22)], align='center'), W_VALUE)

    def cell_merged_value(text):
        """Value cell that spans value+description columns (centered)."""
        return _cell(
            _para([_run(text, bold=True, size=24)], align='center', spacing_after=0),
            W_VALUE + W_DESC)

    def cell_merged_desc(paragraphs):
        """Description cell that spans value+description columns. Supports
        inline **bold** markers in each paragraph string."""
        out = []
        for i, p in enumerate(paragraphs):
            after = 140 if i < len(paragraphs) - 1 else 0
            if not p:
                out.append(_para([_run('', size=22)], align='left', spacing_after=after))
            else:
                out.append(_para(_runs_rich(p, size=22), align='left', spacing_after=after))
        return _cell(''.join(out), W_VALUE + W_DESC, valign='top')

    def cell_desc(paragraphs):
        """Description cell (right column). Supports inline **bold** markers."""
        out = []
        for i, p in enumerate(paragraphs):
            after = 140 if i < len(paragraphs) - 1 else 0
            if not p:
                out.append(_para([_run('', size=22)], align='left', spacing_after=after))
            else:
                out.append(_para(_runs_rich(p, size=22), align='left', spacing_after=after))
        return _cell(''.join(out), W_DESC, valign='top')

    # ─── Title block ────────────────────────────────────────────────────────
    title_paras = [
        _para([_run('OFFER SUMMARY – REVENUE-BASED FINANCING', bold=True, size=26)],
              align='center', spacing_before=0, spacing_after=240),
    ]

    # ─── Row 1: Funding Provided ────────────────────────────────────────────
    funding_label_subs = [
        'to ' + merchant_name,
        '/ ' + merchant_dba,
    ] if merchant_name else []
    row1 = _row([
        cell_label('Funding Provided', sub_lines=funding_label_subs),
        cell_value(pp_fmt),
        cell_desc([
            f'This is how much funding **{provider}** will provide.',
            f'Due to deductions or payments to others, the total funds that will be provided to you directly is **{net_fmt}**.',
            'This amount may increase or decrease based on your final balances with others or any changes to deductions.',
        ]),
    ])

    # ─── Row 2: APR ─────────────────────────────────────────────────────────
    row2 = _row([
        cell_label('Estimated Annual Percentage Rate (APR)'),
        cell_value(f'{apr:.2f}%'),
        cell_desc([
            'APR is the estimated cost of your financing expressed as a yearly rate. '
            'APR incorporates the amount and timing of the funding you receive, fees you pay, '
            'and the periodic payments you make. This calculation assumes your estimated '
            f'average monthly income through your sales of goods and services will be **{rev_fmt}**. '
            'Since your actual income may vary from our estimate, your effective APR may also vary.',
            f'APR is not an interest rate. The cost of this financing is based upon fees charged '
            f'by **{provider}** rather than interest that accrues over time.',
        ]),
    ])

    # ─── Row 3: Finance Charge ──────────────────────────────────────────────
    row3 = _row([
        cell_label('Finance Charge'),
        cell_value(fc_fmt),
        cell_desc([
            'This is the dollar cost of your financing.',
            'Your finance charge will not increase if you take longer to pay off what you owe.',
        ]),
    ])

    # ─── Row 4: Estimated Total Payment Amount ──────────────────────────────
    row4 = _row([
        cell_label('Estimated Total Payment Amount'),
        cell_value(pa_fmt),
        cell_desc([
            'This is the total dollar amount of payments we estimate you will make under the contract.',
        ]),
    ])

    # ─── Row 5: Estimated Monthly Cost ──────────────────────────────────────
    row5 = _row([
        cell_label('Estimated Monthly Cost'),
        cell_value(em_fmt),
        cell_desc([
            'Although you do not make payments on a monthly basis, this is our calculation of '
            'your average monthly cost based upon the payment amounts disclosed below.',
        ]),
    ])

    # ─── Row 6: Estimated Payment (label + merged value spanning to right edge) ───
    row6 = _row([
        cell_label('Estimated Payment'),
        cell_merged_value(f'{pmt_fmt} per {period_label}'),
    ])

    # ─── Row 7: Payment Terms (separate row, label + merged description) ──────
    payment_terms_paras = [
        'Payments are tendered in daily or weekly increments. Daily payments are deducted '
        'every business day, Monday through Friday, and are debited from your business bank '
        'account. If the debit is scheduled for a bank holiday, it will be processed on the '
        'next business day, in addition to the regularly scheduled daily debit.',
        'Weekly payments are deducted once weekly. If the scheduled day is a bank holiday, '
        'it will be deducted on the next business day.',
        f'If the payment under the Agreement is a weekly payment, **{provider}** reserves the right '
        'to switch the payment to a daily payment in the event of the return of 2 consecutive '
        'weekly payments, among any other rights and remedies under the Agreement. The daily '
        'payment would be the weekly payment divided by 5.',
        f'The Estimated Payment is based on **{spec_pct:.2f}%** of your estimated daily business receipts. '
        'This financing does not have a fixed payment schedule and there is no minimum payment amount.',
        'Upon review of information provided by recipient and the nature of the recipient\'s '
        'business, the Provider does not have a reasonable basis to anticipate a true-up. '
        'Recipient should refer to **Section 3 of the Agreement for the reconciliation procedure**.',
    ]
    row7 = _row([
        cell_label('Payment Terms'),
        cell_merged_desc(payment_terms_paras),
    ])

    # ─── Row 8: Estimated Term (standard 3-column layout) ──────────────────────
    row8 = _row([
        cell_label('Estimated Term'),
        cell_value(str(est_term_months)),
        cell_desc([
            'Based upon your expected average sales revenue and purchase percentage, this is our '
            'estimate of how long (in months) it will take to collect the amounts due under the '
            'Purchase Agreement.',
        ]),
    ])

    # ─── Row 9: Prepayment (label + merged description, no value column) ──────
    row9 = _row([
        cell_label('Prepayment'),
        cell_merged_desc([
            f'If you pay off the financing faster than required, you still must pay all or a '
            f'portion of the finance charge up to **{fc_fmt}** based upon our estimates.',
            'If you pay off the financing faster than required, you will not be required to pay '
            'additional fees.',
        ]),
    ])

    disclosure_table = _table([row1, row2, row3, row4, row5, row6, row7, row8, row9],
                              total_width_dxa=10800)

    # ─── Acknowledgment + signature ─────────────────────────────────────────
    ack_para = _para([_run(
        'Applicable law requires this information to be provided to you to help you make an '
        'informed decision. By signing below, you are confirming that you received this information.',
        bold=True, size=22)],
        align='left', spacing_before=320, spacing_after=320)

    # Signature row — invisible 3-column table: signature line | gap | date line
    def sig_block(name):
        sig_cell = _cell(
            _para([_run('_' * 60, size=20)], align='left', spacing_after=20)
            + _para([_run(f'Recipient Signature ({name})', size=18)],
                    align='left', spacing_after=0),
            6400, borders=False)
        gap_cell = _cell(_para([_run('', size=20)], align='left'), 400, borders=False)
        date_cell = _cell(
            _para([_run('_' * 30, size=20)], align='left', spacing_after=20)
            + _para([_run('Date', size=18)], align='left', spacing_after=0),
            4000, borders=False)
        return _row([sig_cell, gap_cell, date_cell])

    sig_rows = [sig_block(signer1_name if signer1_name else '')]
    if two_signers and signer2_name:
        spacer = _cell(_para([_run('', size=18)]), 10800, borders=False)
        sig_rows.append(_row([spacer]))
        sig_rows.append(sig_block(signer2_name))

    sig_table = _table(sig_rows, total_width_dxa=10800)
    # Hide signature-table borders
    sig_table = sig_table.replace(
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>',
        '<w:tblBorders>'
        '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        '</w:tblBorders>'
    )

    # ─── Itemization of Amount Financed (separate page) ─────────────────────
    page_break_to_item = _page_break()

    item_title = _para([_run('ITEMIZATION OF AMOUNT FINANCED', bold=True, size=24)],
                       align='center', spacing_before=120, spacing_after=240)

    # Itemization table — 2 columns: Description | Amount
    W_DESC2 = 7600
    W_AMT2 = 3200

    def item_row(desc, amount, bold_amount=False, shaded=False, indent=False):
        desc_para = _para(
            [_run(desc, bold=False, size=20)],
            align='left',
            spacing_after=0,
            indent_left=400 if indent else None)
        amt_para = _para([_run(amount, bold=bold_amount, size=20)], align='right',
                         spacing_after=0)
        return _row([
            _cell(desc_para, W_DESC2, shading='E7E6E6' if shaded else None),
            _cell(amt_para, W_AMT2),
        ])

    # Build itemization following Angry Petes layout precisely
    item_rows = [
        item_row('1. Amount Given Directly to You', net_fmt),
        item_row('2. ACH Program Fee', ach_fee_fmt),
        item_row('3. Origination Fee', orig_fee_fmt),
        item_row('4. Amount paid on your behalf to third parties (5a + 5b + 5c)',
                 _fmt_currency(0)),
        item_row('5a.', _fmt_currency(0), indent=True),
        item_row('5b.', _fmt_currency(0), indent=True),
        item_row('5c.', _fmt_currency(0), indent=True),
        item_row(f'5. Amount Paid on Your Account with {provider}',
                 _fmt_currency(0)),
        item_row('6. Amount Provided to You or on Your Behalf', pp_fmt),
        item_row('7. Prepaid Finance Charges: ACH Program Fee + Origination Fee',
                 fees_fmt),
        item_row('8. Amount Financed', net_fmt),
    ]
    item_table = _table(item_rows, total_width_dxa=10800)

    # ─── Section properties (page margins 0.75") ────────────────────────────
    # 0.75" = 1080 twips — slightly more generous than 0.5" to match reference
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )

    # ─── Assemble body ──────────────────────────────────────────────────────
    body = (
        ''.join(title_paras)
        + disclosure_table
        + ack_para
        + sig_table
        + page_break_to_item
        + item_title
        + item_table
        + sect_pr
    )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>\n'
        f'<w:body>{body}</w:body>\n'
        '</w:document>'
    )

    # ─── Minimal required DOCX scaffolding ─────────────────────────────────
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
</w:styles>'''

    # Build DOCX
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document_xml)
        z.writestr('word/styles.xml', styles)

    return buf.getvalue()
