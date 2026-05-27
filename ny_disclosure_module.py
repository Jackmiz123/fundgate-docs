"""
NY Commercial Financing Disclosure module — builds the FundGate LLC NY
Offer Summary disclosure as DOCX bytes.

Format mirrors the manually-generated FundGate Olympic reference: three
sections (Offer Summary, Itemization of Amount Financed, Broker
Compensation Disclosure) on three pages.

APR is calculated using the TILA Regulation Z Appendix J actuarial method
with unit-period compounding (52 weekly periods / 260 business-day daily
periods per year). This matches the methodology required by 23 NYCRR
Part 600 (NY DFS Commercial Financing Disclosure Law).

This module is FundGate-only by design — NY disclosure is never branded
as Fundkey. Provider name and email are hardcoded.
"""
import io, zipfile, re
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# Helpers — number / date / dictionary parsing
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
# APR calculation — Reg Z Appendix J actuarial, unit-period compounding
# ─────────────────────────────────────────────────────────────────────────────
def _calculate_apr_ny(purchase_price, payment_amount, frequency, total_purchased):
    """
    Solve for APR using unit-period compounding (NY DFS interpretation).

    purchase_price  — gross funding provided (e.g. $15,000)
    payment_amount  — periodic payment (e.g. $2,248.50 weekly)
    frequency       — 'weekly' or 'daily'
    total_purchased — total dollars to be repaid (e.g. $22,485)

    Returns APR as a percentage (e.g. 422.76 for 422.76%).

    Verified against the FundGate Olympic reference:
      PP=15,000, PMT=2,248.50 weekly, Total=22,485 -> APR = 422.76%
    """
    if purchase_price <= 0 or payment_amount <= 0 or total_purchased <= 0:
        return 0.0

    n_payments = int(round(total_purchased / payment_amount))
    if n_payments <= 0:
        return 0.0

    is_weekly = 'week' in (frequency or '').lower()
    periods_per_year = 52 if is_weekly else 260  # 260 biz days/yr

    # PV equation: PP = PMT × (1 - (1+i)^-N) / i, where i is per-period rate.
    # Solve for i via bisection.
    def pv_at(i):
        if i <= 0:
            return float('inf')
        return payment_amount * (1 - (1 + i) ** (-n_payments)) / i

    lo, hi = 1e-6, 5.0  # bracket: 0.0001% to 500% per period
    if pv_at(hi) > purchase_price:
        # Rate higher than 500%/period — fall through with hi
        return round(hi * periods_per_year * 100, 2)

    mid = lo
    for _ in range(200):
        mid = (lo + hi) / 2
        pv = pv_at(mid)
        if abs(pv - purchase_price) < 0.01:
            break
        if pv > purchase_price:
            lo = mid
        else:
            hi = mid

    apr_pct = mid * periods_per_year * 100
    return round(apr_pct, 2)


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers — build raw OOXML
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
    space = ' xml:space="preserve"' if text != text.strip() else ''
    return f'<w:r>{rpr}<w:t{space}>{_safe(text)}</w:t></w:r>'


def _runs_rich(text, size=22):
    """Parse a string with **bold** inline markers into multiple <w:r> runs."""
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


def _para(runs_xml, align=None, spacing_before=0, spacing_after=0, indent_left=None, keep_next=False):
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
    return ('<w:p><w:r>'
            f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/></w:rPr>'
            '<w:br w:type="page"/></w:r></w:p>')


def _cell(content_xml, width_dxa, vmerge=None, shading=None, borders=True, valign='center'):
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
    tcpr += '<w:tcMar><w:top w:w="160" w:type="dxa"/><w:left w:w="140" w:type="dxa"/>'
    tcpr += '<w:bottom w:w="160" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>'
    tcpr += '</w:tcPr>'
    return f'<w:tc>{tcpr}{content_xml}</w:tc>'


def _row(cells_xml, height_dxa=None):
    tr_pr = ''
    if height_dxa:
        tr_pr = f'<w:trPr><w:trHeight w:val="{height_dxa}" w:hRule="atLeast"/></w:trPr>'
    return f'<w:tr>{tr_pr}{"".join(cells_xml)}</w:tr>'


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
def build_ny_disclosure_bytes(data):
    """
    Build the FundGate NY commercial financing disclosure as DOCX bytes.
    Returns None if state is not NY.

    Three sections, each on its own page:
      Page 1 — OFFER SUMMARY (Offer Summary table + signature block)
      Page 2 — ITEMIZATION OF AMOUNT FINANCED
      Page 3 — BROKER COMPENSATION DISCLOSURE (+ signature block)
    """
    state_code = (data.get('State_of_Organization') or '').upper().strip()
    if state_code != 'NY':
        return None

    # NY disclosure is FundGate-only — hardcoded provider regardless of flags.
    provider = 'FundGate LLC'

    # ── Pull inputs ─────────────────────────────────────────────────────────
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

    avg_monthly_rev = _n(data, 'NY_Avg_Monthly_Revenue')

    ach_freq = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()
    is_weekly = 'week' in ach_freq
    if is_weekly:
        pmt = _n(data, 'Specific_Weekly_Amount')
        period_label = 'week'
    else:
        pmt = _n(data, 'Specific_Daily_Amount')
        period_label = 'business day'

    # APR — NY methodology (unit-period actuarial)
    apr = _calculate_apr_ny(pp, pmt, ach_freq, pa) if (pp > 0 and pa > 0 and pmt > 0) else 0.0

    # Estimated monthly cost: weekly × 52/12  OR  daily × 252/12 (252 biz days/yr)
    if is_weekly:
        est_monthly = round(pmt * 52 / 12, 2)
    else:
        est_monthly = round(pmt * 252 / 12, 2)

    # Estimated term (months) — Total Payments / (Avg Monthly Revenue × Spec %)
    monthly_capture = avg_monthly_rev * spec_pct / 100.0
    if monthly_capture > 0 and pa > 0:
        # Round up (ceiling) to next whole month, matching reference
        from math import ceil
        est_term_months = max(1, ceil(pa / monthly_capture - 0.0001))
    elif pmt > 0 and pa > 0:
        # Fallback if revenue not provided — derive from periodic count
        n_payments = int(round(pa / pmt))
        from math import ceil
        est_term_months = max(1, ceil(n_payments / (52/12) if is_weekly else n_payments / 21))
    else:
        est_term_months = 0

    two_signers = bool(data.get('twoSigners', False))
    signer1_name = (data.get('Owner_Guarantor_1', '') or '').upper()
    signer1_title = (data.get('Title', '') or '').upper() or 'MEMBER'
    signer2_name = (data.get('Owner_Guarantor_2', '') or '').upper() if two_signers else ''
    signer2_title = (data.get('Title_2', '') or '').upper() if two_signers else ''
    if two_signers and not signer2_title:
        signer2_title = 'MEMBER'

    # Display formats
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

    # ── Column widths — total 10800 dxa ─────────────────────────────────────
    W_LABEL = 2400
    W_VALUE = 2400
    W_DESC  = 6000

    # ── Cell builders ───────────────────────────────────────────────────────
    def cell_label(text, sub_lines=None):
        paras = [_para([_run(text, bold=True, size=22)], align='left', spacing_after=0)]
        if sub_lines:
            paras.append(_para([_run('', size=18)], align='left', spacing_after=0))
            for s in sub_lines:
                paras.append(_para([_run(s, bold=False, size=20)],
                                   align='left', spacing_after=0))
        return _cell(''.join(paras), W_LABEL, valign='top')

    def cell_value(text, sub=None):
        paras = [_para([_run(text, bold=True, size=24)], align='center', spacing_after=0)]
        if sub:
            paras.append(_para([_run(sub, size=20)], align='center', spacing_after=0))
        return _cell(''.join(paras), W_VALUE, valign='center')

    def cell_merged_value(text):
        """Value cell that spans across the Value+Desc columns (for rows with no description)."""
        para = _para([_run(text, bold=True, size=24)], align='center', spacing_after=0)
        merged_width = W_VALUE + W_DESC
        cell_xml = '<w:tc>'
        cell_xml += '<w:tcPr>'
        cell_xml += f'<w:tcW w:w="{merged_width}" w:type="dxa"/>'
        cell_xml += '<w:gridSpan w:val="2"/>'
        cell_xml += ('<w:tcBorders>'
                     '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '</w:tcBorders>')
        cell_xml += '<w:vAlign w:val="center"/>'
        cell_xml += '<w:tcMar><w:top w:w="160" w:type="dxa"/><w:left w:w="140" w:type="dxa"/>'
        cell_xml += '<w:bottom w:w="160" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>'
        cell_xml += '</w:tcPr>'
        cell_xml += para
        cell_xml += '</w:tc>'
        return cell_xml

    def cell_merged_desc(paragraphs):
        """Description cell that spans Value+Desc columns (for rows with no value, just description)."""
        out = []
        for idx, p in enumerate(paragraphs):
            after = 120 if idx < len(paragraphs) - 1 else 0
            if not p:
                out.append(_para([_run('', size=22)], align='left', spacing_after=after))
            else:
                out.append(_para(_runs_rich(p, size=22), align='left', spacing_after=after))
        merged_width = W_VALUE + W_DESC
        cell_xml = '<w:tc>'
        cell_xml += '<w:tcPr>'
        cell_xml += f'<w:tcW w:w="{merged_width}" w:type="dxa"/>'
        cell_xml += '<w:gridSpan w:val="2"/>'
        cell_xml += ('<w:tcBorders>'
                     '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     '</w:tcBorders>')
        cell_xml += '<w:vAlign w:val="top"/>'
        cell_xml += '<w:tcMar><w:top w:w="160" w:type="dxa"/><w:left w:w="140" w:type="dxa"/>'
        cell_xml += '<w:bottom w:w="160" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar>'
        cell_xml += '</w:tcPr>'
        cell_xml += ''.join(out)
        cell_xml += '</w:tc>'
        return cell_xml

    def cell_desc(paragraphs):
        out = []
        for idx, p in enumerate(paragraphs):
            after = 120 if idx < len(paragraphs) - 1 else 0
            if not p:
                out.append(_para([_run('', size=22)], align='left', spacing_after=after))
            else:
                out.append(_para(_runs_rich(p, size=22), align='left', spacing_after=after))
        return _cell(''.join(out), W_DESC, valign='top')

    # ─── Title ─────────────────────────────────────────────────────────────
    title_paras = [
        _para([_run('OFFER SUMMARY \u2013 REVENUE-BASED FINANCING', bold=True, size=26)],
              align='center', spacing_before=0, spacing_after=240),
    ]

    # ─── Row 1: Funding Provided ───────────────────────────────────────────
    funding_label_subs = []
    if merchant_name:
        funding_label_subs.append(f'to {merchant_name}')
    row1 = _row([
        cell_label('Funding Provided', sub_lines=funding_label_subs),
        cell_value(pp_fmt, sub='Amount Financed'),
        cell_desc([
            f'This is how much funding **{provider}** will provide.',
            f'Due to deductions or payments to others, the total funds that will be provided to you directly is **{net_fmt}**.',
            'This amount may increase or decrease based on your final balances with others or any changes to deductions.',
        ]),
    ], height_dxa=2400)

    # ─── Row 2: Estimated APR ──────────────────────────────────────────────
    row2 = _row([
        cell_label('Estimated Annual Percentage Rate (APR)'),
        cell_value(f'{apr:.2f}%'),
        cell_desc([
            'APR is the estimated cost of your financing expressed as a yearly rate. '
            'APR incorporates the amount and timing of the funding you receive, finance '
            'charges you pay, and the periodic payments you make. This calculation '
            f'assumes your estimated average monthly income through your sales of goods '
            f'and services will be **{rev_fmt}**.',
            'Since your actual income may vary from our estimate, your effective APR '
            f'may also vary. APR is not an interest rate. The cost of this financing is '
            f'based upon fees charged by **{provider}** rather than interest that accrues '
            f'over time.',
        ]),
    ], height_dxa=3400)

    # ─── Row 3: Finance Charge ─────────────────────────────────────────────
    row3 = _row([
        cell_label('Finance Charge'),
        cell_value(fc_fmt),
        cell_desc([
            'This is the dollar cost of your financing. Your finance charge will not '
            'increase if you take longer to pay off what you owe.',
        ]),
    ], height_dxa=1700)

    # ─── Row 4: Estimated Total Payment Amount ────────────────────────────
    row4 = _row([
        cell_label('Estimated Total Payment Amount'),
        cell_value(pa_fmt),
        cell_desc([
            'This is the total dollar amount of payments we estimate you will make under '
            'the contract.',
        ]),
    ], height_dxa=1700)

    # ─── Row 5: Estimated Monthly Cost ────────────────────────────────────
    row5 = _row([
        cell_label('Estimated Monthly Cost'),
        cell_value(em_fmt),
        cell_desc([
            'Although you do not make payments on a monthly basis, this is our '
            'calculation of your average monthly cost based upon the payment amounts '
            'disclosed below.',
        ]),
    ], height_dxa=1900)

    # ─── Row 6: Estimated Payment (label + merged value spanning to right) ───
    row6 = _row([
        cell_label('Estimated Payment'),
        cell_merged_value(f'{pmt_fmt} per {period_label}'),
    ], height_dxa=1700)

    # ─── Row 7: Estimated Term ────────────────────────────────────────────
    row7 = _row([
        cell_label('Estimated Term'),
        cell_value(str(est_term_months)),
        cell_desc([
            'Based upon your expected average sales revenue and purchased percentage, '
            'this is our estimate of how long it will take to collect the amounts due '
            'under the Purchase Agreement.',
        ]),
    ])

    # ─── Row 8: Prepayment ─────────────────────────────────────────────────
    row8 = _row([
        cell_label('Prepayment'),
        cell_merged_desc([
            f'If you pay off the financing faster than required, you still must pay all '
            f'or a portion of the finance charge up to **{fc_fmt}** based upon our estimates.',
            'If you pay off the financing faster than required, you will not be required '
            'to pay additional fees.',
        ]),
    ])

    # ─── Row 9: Collateral Requirements ────────────────────────────────────
    row9 = _row([
        cell_label('Collateral Requirements'),
        cell_merged_desc([
            f'**{provider}** has a security interest in the Future Receipts of the Merchant. '
            'Future Receipts includes all payments made by cash, check, ACH, or other '
            'electronic transfer, credit card, debit card, bank card, charge card or other '
            'form of monetary payment in the ordinary course of Merchant\u2019s business, '
            'accounts and payment intangibles, and all proceeds and products of the foregoing.',
        ]),
    ])

    disclosure_table_p1 = _table(
        [row1, row2, row3, row4, row5, row6],
        total_width_dxa=10800,
    )
    disclosure_table_p2 = _table(
        [row7, row8, row9],
        total_width_dxa=10800,
    )

    # ─── Acknowledgment paragraph ──────────────────────────────────────────
    ack_para = _para([_run(
        'Applicable law requires this information to be provided to you to help you make '
        'an informed decision. By signing below, you are confirming that you received '
        'this information.',
        bold=True, size=22)],
        align='left', spacing_before=320, spacing_after=320)

    # ─── Signature block (page 1) ──────────────────────────────────────────
    def sig_block(name, title):
        label = f'{name} - {title}' if (name and title) else (name or '')
        sig_cell = _cell(
            _para([_run('_' * 60, size=20)], align='left', spacing_after=20)
            + _para([_run(label, size=18)], align='left', spacing_after=0),
            6400, borders=False)
        gap_cell = _cell(_para([_run('', size=20)], align='left'), 400, borders=False)
        date_cell = _cell(
            _para([_run('_' * 30, size=20)], align='left', spacing_after=20)
            + _para([_run('Date', size=18)], align='left', spacing_after=0),
            4000, borders=False)
        return _row([sig_cell, gap_cell, date_cell])

    sig_rows = [sig_block(signer1_name, signer1_title)]
    if two_signers and signer2_name:
        spacer = _cell(_para([_run('', size=18)]), 10800, borders=False)
        sig_rows.append(_row([spacer]))
        sig_rows.append(sig_block(signer2_name, signer2_title))

    sig_table_offer = _build_borderless_table(sig_rows)

    # ─── Page 2: Itemization of Amount Financed ────────────────────────────
    item_title = _para([_run('ITEMIZATION OF AMOUNT FINANCED', bold=True, size=24)],
                       align='center', spacing_before=120, spacing_after=240)

    W_ITEM_DESC = 8000
    W_ITEM_AMT = 2800

    def item_row(desc, amount, bold_amount=False, indent=False):
        desc_para = _para(
            [_run(desc, bold=False, size=20)],
            align='left', spacing_after=0,
            indent_left=400 if indent else None)
        amt_para = _para([_run(amount, bold=bold_amount, size=20)],
                         align='right', spacing_after=0)
        return _row([
            _cell(desc_para, W_ITEM_DESC, valign='center'),
            _cell(amt_para, W_ITEM_AMT, valign='center'),
        ])

    item_rows = [
        item_row('1.  Amount Given Directly to You', net_fmt),
        item_row('2.  ACH Program Fee', ach_fee_fmt),
        item_row('3.  Origination Fee', orig_fee_fmt),
        item_row(f'4.  Amount Paid on Your Account with {provider} Advance #',
                 _fmt_currency(0)),
        item_row('5.  Amount paid on your behalf to third parties (5a + 5b + 5c)',
                 _fmt_currency(0)),
        item_row('5a.', _fmt_currency(0), indent=True),
        item_row('5b.', _fmt_currency(0), indent=True),
        item_row('5c.', _fmt_currency(0), indent=True),
        item_row('6.  Amount Provided to You or on Your Behalf (1 + 2 + 3 + 4 + 5)', pp_fmt),
        item_row('7.  Prepaid Finance Charges: ACH Program Fee + Origination Fee', fees_fmt),
        item_row('8.  Amount Financed (6 minus 7)', net_fmt, bold_amount=True),
    ]
    item_table = _table(item_rows, total_width_dxa=10800)

    # ─── Page 3: Broker Compensation Disclosure ────────────────────────────
    broker_title = _para([_run('BROKER COMPENSATION DISCLOSURE', bold=True, size=24)],
                         align='center', spacing_before=120, spacing_after=240)

    broker_body = _para(
        _runs_rich(
            f'In connection with this proposed commercial financing transaction, '
            f'**{provider}** will pay compensation directly to a broker for the broker\u2019s '
            f'role in the transaction. If a broker charges you directly for a separate '
            f'broker fee, that fee is in addition to the broker compensation we will pay '
            f'to the broker.',
            size=22),
        align='left', spacing_before=0, spacing_after=240,
    )

    broker_ack = _para([_run(
        'Applicable law requires this information to be provided to you to help you make '
        'an informed decision. By signing below, you are confirming that you received '
        'this information.',
        bold=True, size=22)],
        align='left', spacing_before=120, spacing_after=320)

    broker_sig_rows = [sig_block(signer1_name, signer1_title)]
    if two_signers and signer2_name:
        spacer = _cell(_para([_run('', size=18)]), 10800, borders=False)
        broker_sig_rows.append(_row([spacer]))
        broker_sig_rows.append(sig_block(signer2_name, signer2_title))

    sig_table_broker = _build_borderless_table(broker_sig_rows)

    # ─── Section properties (page margins ~0.75") ─────────────────────────
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
        + disclosure_table_p1
        + _page_break()
        + disclosure_table_p2
        + ack_para
        + sig_table_offer
        + _page_break()
        + item_title
        + item_table
        + _page_break()
        + broker_title
        + broker_body
        + broker_ack
        + sig_table_broker
        + sect_pr
    )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>\n'
        f'<w:body>{body}</w:body>\n'
        '</w:document>'
    )

    # ─── DOCX scaffolding ──────────────────────────────────────────────────
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '</Types>')

    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')

    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '</Relationships>')

    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults>'
              '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
              '</w:docDefaults>'
              '</w:styles>')

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document_xml)
        z.writestr('word/styles.xml', styles)

    return buf.getvalue()


def _build_borderless_table(rows):
    """Helper: build a table with all borders set to nil (used for signature rows)."""
    tbl = _table(rows, total_width_dxa=10800)
    return tbl.replace(
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
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
