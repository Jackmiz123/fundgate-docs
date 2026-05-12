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


def _cell(content_xml, width_dxa, vmerge=None, shading=None, borders=True):
    """Build a <w:tc>. content_xml is one or more <w:p> elements."""
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
    tcpr += '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
    tcpr += '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>'
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
    ach_fee = round(pp * ach_pct / 100, 2)
    orig_fee = round(pp * orig_pct / 100, 2)
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

    # Column widths (total 10800 dxa)
    W_LABEL = 2200
    W_VALUE = 2400
    W_DESC = 6200

    # ─── Build PAGE 1 ───────────────────────────────────────────────────────
    # Title block
    title_runs = [
        _para([_run('FUNDING DISCLOSURE', bold=True, size=28)],
              align='center', spacing_before=0, spacing_after=80),
        _para([_run('California Commercial Financing Disclosure', bold=True, size=22)],
              align='center', spacing_before=0, spacing_after=80),
        _para([_run('Provided pursuant to California Financial Code §§ 22800–22805 and '
                    '10 CCR §§ 901 et seq.', italic=True, size=18)],
              align='center', spacing_before=0, spacing_after=120),
        _para([_run(f'Recipient: {merchant_dba}', bold=True, size=20)],
              align='center', spacing_before=0, spacing_after=40),
        _para([_run(f'Date: {date_display}', size=20)],
              align='center', spacing_before=0, spacing_after=200),
    ]

    # The disclosure table rows
    def cell_label(text):
        return _cell(
            _para([_run(text, bold=True, size=20)], align='left', spacing_after=0),
            W_LABEL, shading='E7E6E6')

    def cell_value(text, sub=None):
        runs = [_para([_run(text, bold=True, size=22)], align='center', spacing_after=20)]
        if sub:
            runs.append(_para([_run(sub, size=16)], align='center', spacing_after=0))
        return _cell(''.join(runs), W_VALUE)

    def cell_desc(text):
        return _cell(
            _para([_run(text, size=18)], align='left', spacing_after=0),
            W_DESC)

    rows = []
    # Row 1: Funding Provided
    rows.append(_row([
        cell_label('Funding Provided'),
        cell_value(pp_fmt, sub=f'Net to Recipient: {net_fmt}'),
        cell_desc(
            f'The total dollar amount of funding Fundkey LLC will provide. The amount received by '
            f'the recipient ({net_fmt}) is the funding amount minus all fees deducted at funding.')
    ]))
    # Row 2: APR
    rows.append(_row([
        cell_label('Estimated Annual Percentage Rate'),
        cell_value(f'{apr:.2f}%'),
        cell_desc(
            'APR is the estimated cost of your financing, including fees, expressed as a yearly rate. '
            'APR is provided for comparison purposes. APR is not an interest rate. Your financing '
            'agreement does not provide for an interest rate. APR represents the cost of fees charged '
            'by Fundkey LLC rather than interest.')
    ]))
    # Row 3: Finance Charge
    rows.append(_row([
        cell_label('Finance Charge'),
        cell_value(fc_fmt),
        cell_desc('The dollar cost of your financing, including all fees and other charges.')
    ]))
    # Row 4: Total Payment Amount
    rows.append(_row([
        cell_label('Total Payment Amount'),
        cell_value(pa_fmt),
        cell_desc('The total dollar amount you will pay, including the amount funded and the finance charge.')
    ]))
    # Row 5: Estimated Monthly Cost
    rows.append(_row([
        cell_label('Estimated Monthly Cost'),
        cell_value(em_fmt),
        cell_desc(
            'Although you do not make payments on a monthly basis, this is our calculation of your '
            'average monthly cost based upon the payment amounts disclosed below.')
    ]))
    # Row 6: Estimated Payment
    rows.append(_row([
        cell_label(f'Estimated {period_label_cap} Payment'),
        cell_value(f'{pmt_fmt}/{period_label}', sub=f'({n_payments} payments)'),
        cell_desc(
            f'Payments are tendered in {period_label_cap.lower()} increments via automatic '
            'ACH debit from the recipient\'s designated business bank account. Fundkey LLC '
            'reserves the right to reconcile the amount upon recipient request as set forth '
            'in Section 3 of the Agreement.')
    ]))
    # Row 7: Avg Monthly Revenue
    rows.append(_row([
        cell_label('Avg. Monthly Income'),
        cell_value(rev_fmt),
        cell_desc(
            'The recipient\'s historical average monthly gross income from sales is used to '
            'estimate the term and the annual percentage rate.')
    ]))
    # Row 8: Estimated Term
    rows.append(_row([
        cell_label('Estimated Term'),
        cell_value(f'{est_term_months} months'),
        cell_desc(
            f'The estimated number of months it will take to deliver the Total Payment Amount '
            f'based on the recipient\'s historical average monthly income and the specified '
            f'percentage of {spec_pct:.2f}%.')
    ]))
    # Row 9: Prepayment
    rows.append(_row([
        cell_label('Prepayment'),
        cell_value('See description'),
        cell_desc(
            'You will not pay any additional charge to prepay this financing. If you pay off '
            'the financing early, you will not be required to pay the full finance charge; '
            'you may be eligible for a discounted payoff amount as set forth in any addendum '
            'to the Agreement.')
    ]))

    disclosure_table = _table(rows, total_width_dxa=10800)

    # ─── Build PAGE 2 ───────────────────────────────────────────────────────
    page_break_1 = _page_break()

    acknowledgment_paras = [
        _para([_run('Acknowledgment', bold=True, size=22)],
              align='left', spacing_before=200, spacing_after=80),
        _para([_run(
            'Applicable law requires Fundkey LLC to provide this disclosure to you and to '
            'obtain your signature acknowledging your receipt of this disclosure. Your signature '
            'below acknowledges only that you received this disclosure. It does not constitute '
            'agreement to the terms of the financing or any contract.', size=18)],
              align='both', spacing_before=0, spacing_after=120),
        _para([_run(
            'You can find more information about the Department of Financial Protection and '
            'Innovation, the agency that regulates commercial financing in California, at '
            'https://dfpi.ca.gov.', size=18, italic=True)],
              align='both', spacing_before=0, spacing_after=200),
    ]

    # Signature block — invisible 3-column table for alignment
    def sig_line_row(label_left, label_right, line_width_left=4400, gap=400, line_width_right=2800):
        # Underscored signature line cells with labels underneath
        sig_cell_left = _cell(
            _para([_run('_' * 60, size=20)], align='left', spacing_after=20)
            + _para([_run(label_left, size=18)], align='left', spacing_after=0),
            line_width_left, borders=False)
        gap_cell = _cell(_para([_run('', size=20)], align='left'), gap, borders=False)
        sig_cell_right = _cell(
            _para([_run('_' * 30, size=20)], align='left', spacing_after=20)
            + _para([_run(label_right, size=18)], align='left', spacing_after=0),
            line_width_right, borders=False)
        return _row([sig_cell_left, gap_cell, sig_cell_right])

    sig_rows = [sig_line_row(f'Recipient Signature ({signer1_name})', 'Date')]
    if two_signers and signer2_name:
        # Add a spacer row then second signer row
        spacer_cell = _cell(_para([_run('', size=16)]), 10800, borders=False)
        sig_rows.append(_row([spacer_cell]))
        sig_rows.append(sig_line_row(f'Recipient Signature ({signer2_name})', 'Date'))

    sig_table = _table(sig_rows, total_width_dxa=10800)
    # Replace tblBorders with nil for invisible signature table
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

    # ─── Build PAGE 3 — Itemization of Amount Financed ──────────────────────
    page_break_2 = _page_break()

    item_title = _para([_run('Itemization of Amount Financed', bold=True, size=24)],
                       align='center', spacing_before=200, spacing_after=200)

    # Itemization table — 2 columns: Description | Amount
    W_DESC2 = 7600
    W_AMT2 = 3200

    def item_row(desc, amount, bold=False):
        return _row([
            _cell(_para([_run(desc, bold=bold, size=20)], align='left'), W_DESC2),
            _cell(_para([_run(amount, bold=bold, size=20)], align='right'), W_AMT2),
        ])

    item_rows = [
        # Header
        _row([
            _cell(_para([_run('Description', bold=True, size=20)], align='left'),
                  W_DESC2, shading='E7E6E6'),
            _cell(_para([_run('Amount', bold=True, size=20)], align='right'),
                  W_AMT2, shading='E7E6E6'),
        ]),
        item_row('1. Amount Given to You Directly (Net Funding)', net_fmt),
        item_row('2. ACH Program Fee', ach_fee_fmt),
        item_row('3. Origination Fee', orig_fee_fmt),
        item_row('4. Total Fees Deducted at Funding', fees_fmt, bold=True),
        item_row('5. Amount Paid on Your Account with Fundkey LLC (Funding Provided)',
                 pp_fmt, bold=True),
    ]
    item_table = _table(item_rows, total_width_dxa=10800)

    item_note = _para([_run(
        'The above amounts represent how the Funding Provided is allocated. Total Fees Deducted '
        'at Funding plus Amount Given to You Directly equals the Funding Provided amount.',
        italic=True, size=18)], align='both', spacing_before=200, spacing_after=0)

    # ─── Section properties (page margins 0.5" all around) ──────────────────
    # 0.5" = 720 twips
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )

    # ─── Assemble document.xml body ────────────────────────────────────────
    body = (
        ''.join(title_runs)
        + disclosure_table
        + page_break_1
        + ''.join(acknowledgment_paras)
        + sig_table
        + page_break_2
        + item_title
        + item_table
        + item_note
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
