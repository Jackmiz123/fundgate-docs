"""
CT Commercial (Sales-Based) Financing Disclosure module — builds the
FundGate LLC CT disclosure as DOCX bytes.

Format mirrors the manually-generated reference disclosure (LB & O / BizFund
sample): two pages.

Page 1 — Single table with split header (left: financial details,
         right: provider/recipient info), payment schedule checkboxes,
         description rows, broker compensation, prepayment, signature line.

Page 2 — Narrative continuation: variable payment method, fees not in
         finance charge, renewal info, statutory boilerplate, signature.

This module is FundGate-only by design — CT disclosure is never branded
as Fundkey. Provider name, address, phone, and email are hardcoded.

Required by Part XVI of Chapter 669 of the 2024 Supplement to the
Connecticut General Statutes.
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
        return str(val) if val else '$0.00'


def _fmt_date_dash(val):
    """CT format uses MM-DD-YYYY with dashes (e.g. 11-10-2025)."""
    if not val:
        return ''
    for fmt in ('%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d'):
        try:
            return datetime.strptime(str(val).strip(), fmt).strftime('%m-%d-%Y')
        except Exception:
            pass
    return str(val)


def _n(data, key, default=0.0):
    try:
        return float(str(data.get(key, default)).replace('$', '').replace(',', '').replace('%', ''))
    except Exception:
        return default


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers
# ─────────────────────────────────────────────────────────────────────────────
FONT = 'Times New Roman'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def _safe(s):
    return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _run(text, bold=False, italic=False, size=20, underline=False):
    rpr = '<w:rPr>'
    rpr += f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
    if bold:
        rpr += '<w:b/><w:bCs/>'
    if italic:
        rpr += '<w:i/><w:iCs/>'
    if underline:
        rpr += '<w:u w:val="single"/>'
    rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    rpr += '</w:rPr>'
    space = ' xml:space="preserve"' if (text != text.strip() or '  ' in text) else ''
    return f'<w:r>{rpr}<w:t{space}>{_safe(text)}</w:t></w:r>'


def _runs_rich(text, size=20):
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


def _para(runs_xml, align=None, spacing_before=0, spacing_after=0, indent_left=None):
    ppr = '<w:pPr>'
    if align:
        ppr += f'<w:jc w:val="{align}"/>'
    if spacing_before or spacing_after:
        ppr += f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}"/>'
    if indent_left is not None:
        ppr += f'<w:ind w:left="{indent_left}"/>'
    ppr += '</w:pPr>'
    if isinstance(runs_xml, (list, tuple)):
        runs_xml = ''.join(runs_xml)
    return f'<w:p>{ppr}{runs_xml}</w:p>'


def _page_break():
    return ('<w:p><w:r>'
            f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/></w:rPr>'
            '<w:br w:type="page"/></w:r></w:p>')


def _cell(content_xml, width_dxa, gridspan=None, valign='top', vmerge=None):
    tcpr = '<w:tcPr>'
    tcpr += f'<w:tcW w:w="{width_dxa}" w:type="dxa"/>'
    if gridspan:
        tcpr += f'<w:gridSpan w:val="{gridspan}"/>'
    if vmerge:
        tcpr += f'<w:vMerge w:val="{vmerge}"/>'
    tcpr += ('<w:tcBorders>'
             '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '</w:tcBorders>')
    if valign:
        tcpr += f'<w:vAlign w:val="{valign}"/>'
    tcpr += '<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
    tcpr += '<w:bottom w:w="100" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'
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
# PROVIDER constants — hardcoded for FundGate LLC. Update the phone number
# when Jack provides the live FundGate phone (currently placeholder).
PROVIDER_NAME    = 'FundGate LLC'
PROVIDER_ADDR    = '1202 Avenue U, Suite 1175, Brooklyn, NY 11229'
PROVIDER_PHONE   = '929-256-7464'
PROVIDER_EMAIL   = 'admin@fundgatellc.com'


def build_ct_disclosure_bytes(data):
    """
    Build the FundGate CT commercial financing disclosure as DOCX bytes.
    Returns None if state is not CT.
    """
    state_code = (data.get('State_of_Organization') or '').upper().strip()
    if state_code != 'CT':
        return None

    # ── Pull inputs ─────────────────────────────────────────────────────────
    # Recipient name: when DBA differs from legal name, show "LEGAL DBA DBA".
    # Otherwise just legal name. Matches contract DBA handling logic.
    _legal = (data.get('Merchant_Legal_Name', '') or '').strip()
    _dba   = (data.get('Merchant_DBA', '') or '').strip()
    if _dba and _dba.upper() != _legal.upper():
        merchant_name = f'{_legal} DBA {_dba}'.upper()
    else:
        merchant_name = _legal.upper()
    merchant_addr = (data.get('Executive_Office_Address', '') or '').upper()
    disclosure_date = _fmt_date_dash(data.get('Agreement_Date', ''))

    pp = _n(data, 'Purchase_Price')
    pa = _n(data, 'Purchased_Amount')
    ach_pct = _n(data, 'ACH_Program_Fee_Percentage')
    orig_pct = _n(data, 'Origination_Fee_Percentage')
    # Support both % and $ modes for fees (v8+ feature).
    ach_fee_mode = (data.get('ACH_Program_Fee_Mode', 'pct') or 'pct').lower()
    orig_fee_mode = (data.get('Origination_Fee_Mode', 'pct') or 'pct').lower()
    ach_fee = ach_pct if ach_fee_mode == 'dollar' else round(pp * ach_pct / 100, 2)
    orig_fee = orig_pct if orig_fee_mode == 'dollar' else round(pp * orig_pct / 100, 2)
    total_fees = round(ach_fee + orig_fee, 2)
    disbursement = round(pp - total_fees, 2)
    finance_charge = round(pa - disbursement, 2)

    # Specified Percentage and frequency
    spec_pct_raw = data.get('Specified_Percentage', '0')
    try:
        spec_pct = float(str(spec_pct_raw).replace('%', '').replace(',', ''))
    except Exception:
        spec_pct = 0.0

    ach_freq = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()
    is_weekly = 'week' in ach_freq
    if is_weekly:
        pmt = _n(data, 'Specific_Weekly_Amount')
        freq_label_pg1 = 'Weekly'
        freq_label_pg2_checked = 'weekly'
    else:
        pmt = _n(data, 'Specific_Daily_Amount')
        freq_label_pg1 = 'Daily (Business Day)'
        freq_label_pg2_checked = 'daily'

    # Estimated Time Period — expressed in DAYS per CT statute.
    # Payment_Term_Count is the number of PAYMENTS the user enters on the form
    # (number of weeks for weekly deals, number of business days for daily deals).
    # To convert to total days:
    #   Weekly: each weekly payment spans 5 business days → multiply by 5
    #     (e.g. 44 weekly payments = 220 days, matching the BizFund/LB&O reference)
    #   Daily:  each daily payment IS 1 business day → use as-is
    #     (e.g. 100 daily payments = 100 days)
    term_count_raw = data.get('Payment_Term_Count', '') or ''
    try:
        num_payments = int(float(str(term_count_raw).replace(',', '').strip()))
    except Exception:
        # Fallback: derive from total payment / per-payment amount
        if pmt > 0 and pa > 0:
            num_payments = int(round(pa / pmt))
        else:
            num_payments = 0
    term_count = num_payments * 5 if is_weekly else num_payments

    # Average projected monthly payments
    if is_weekly:
        avg_monthly = round(pmt * 52 / 12, 2)
    else:
        avg_monthly = round(pmt * 252 / 12, 2)  # 252 business days/year

    # Broker compensation (new in v10 form input)
    broker_pays = (data.get('CT_Broker_Pays', '') or '').strip().lower() in ('yes', 'true', '1', 'y')
    broker_amount = _n(data, 'CT_Broker_Amount')

    two_signers = bool(data.get('twoSigners', False))
    signer1_name = (data.get('Owner_Guarantor_1', '') or '').upper()
    signer2_name = (data.get('Owner_Guarantor_2', '') or '').upper() if two_signers else ''

    # Formatted strings
    pp_fmt = _fmt_currency(pp)
    pa_fmt = _fmt_currency(pa)
    fees_fmt = _fmt_currency(total_fees)
    disb_fmt = _fmt_currency(disbursement)
    fc_fmt = _fmt_currency(finance_charge)
    pmt_fmt = _fmt_currency(pmt)
    avg_monthly_fmt = _fmt_currency(avg_monthly)
    broker_amount_fmt = _fmt_currency(broker_amount) if broker_pays else '$0.00'

    # ── Title paragraphs ────────────────────────────────────────────────────
    title1 = _para([_run('COMMERCIAL (SALES-BASED) FINANCING DISCLOSURE FORM \u2013 PAGE 1',
                         bold=True, size=24)],
                   align='center', spacing_after=40)
    title2 = _para([_run('(pursuant to Part XVI of Chapter 669 of the 2024 Supplement to the General Statutes)',
                         size=20)],
                   align='center', spacing_after=200)

    # ── Page 1 — Main table ─────────────────────────────────────────────────
    # Layout: 2-column rows (label-value) for left half, OR 4-column rows
    # with right half containing label-value too. Match BizFund: top header
    # block has split L/R columns; main rows span full width with label|value.
    #
    # Column widths (total 10800):
    LEFT_LBL_W = 2900   # label
    LEFT_VAL_W = 2400   # value
    RIGHT_LBL_W = 2300  # label
    RIGHT_VAL_W = 3200  # value

    def label_cell(text, width, bold=True, gridspan=None, sub=None):
        paras = [_para([_run(text, bold=bold, size=20)], align='left', spacing_after=0)]
        if sub:
            paras.append(_para([_run(sub, italic=False, size=16)], align='left', spacing_after=0))
        return _cell(''.join(paras), width, gridspan=gridspan, valign='top')

    def value_cell(text, width, bold=True, gridspan=None, align='left'):
        return _cell(_para([_run(text, bold=bold, size=20)], align=align, spacing_after=0),
                     width, gridspan=gridspan, valign='center')

    def kv_pair_cell(label, value, width, gridspan=None):
        """A cell with bold label inline followed by plain value."""
        runs = _run(label, bold=True, size=20) + _run(' ' + value, bold=False, size=20)
        return _cell(_para(runs, align='left', spacing_after=0), width, gridspan=gridspan, valign='top')

    # Right column: ONE merged box containing all provider/recipient info
    # stacked vertically. Uses Word's vMerge — first row holds the content
    # with vMerge="restart", subsequent rows have empty cells with
    # vMerge="continue" so the box visually spans rows 1–6.
    right_box_content = (
        _para(_run('Disclosure Date:', bold=True, size=20) + _run(' ' + disclosure_date, size=20),
              align='left', spacing_after=200) +
        _para(_run("Recipient's Name:", bold=True, size=20) + _run(' ' + merchant_name, size=20),
              align='left', spacing_after=200) +
        _para(_run("Recipient's Address:", bold=True, size=20) + _run(' ' + merchant_addr, size=20),
              align='left', spacing_after=200) +
        _para(_run("Provider's Name:", bold=True, size=20) + _run(' ' + PROVIDER_NAME, size=20),
              align='left', spacing_after=200) +
        _para(_run("Provider's Address:", bold=True, size=20) + _run(' ' + PROVIDER_ADDR, size=20),
              align='left', spacing_after=200) +
        _para(_run("Provider's Phone Number:", bold=True, size=20) + _run(' ' + PROVIDER_PHONE, size=20),
              align='left', spacing_after=60) +
        _para(_run("Provider's E-mail Address:", bold=True, size=20) + _run(' ' + PROVIDER_EMAIL, size=20),
              align='left', spacing_after=0)
    )
    RIGHT_W = RIGHT_LBL_W + RIGHT_VAL_W

    def right_cell_restart():
        # First row holds all the content
        return _cell(right_box_content, RIGHT_W, gridspan=2, valign='top', vmerge='restart')

    def right_cell_continue():
        # Subsequent rows show as part of the same visual cell (no content)
        return _cell(_para([_run('', size=20)], align='left'),
                     RIGHT_W, gridspan=2, valign='top', vmerge='continue')

    # Row 1: Total Amount of Commercial Financing | (right box starts)
    row1 = _row([
        label_cell('Total Amount of the Commercial Financing', LEFT_LBL_W),
        value_cell(pp_fmt, LEFT_VAL_W),
        right_cell_restart(),
    ])

    # Row 2: Fees Deducted | (right box continues)
    row2 = _row([
        label_cell('Fees Deducted or Withheld at Disbursement', LEFT_LBL_W),
        value_cell(fees_fmt, LEFT_VAL_W),
        right_cell_continue(),
    ])

    # Row 3: Disbursement Amount | (right box continues)
    disb_sub = ('* Amount Paid to Recipient or on the Recipient\u2019s Behalf, Excluding Finance '
                'Charges Deducted or Withheld at Disbursement')
    disb_label_para = (
        _para([_run('Disbursement Amount', bold=True, size=20)], align='left', spacing_after=20) +
        _para([_run(disb_sub, italic=True, size=15)], align='left', spacing_after=0)
    )
    row3 = _row([
        _cell(disb_label_para, LEFT_LBL_W, valign='top'),
        value_cell(disb_fmt, LEFT_VAL_W),
        right_cell_continue(),
    ])

    # Row 4: Finance Charge | (right box continues)
    row4 = _row([
        label_cell('Finance Charge', LEFT_LBL_W),
        value_cell(fc_fmt, LEFT_VAL_W),
        right_cell_continue(),
    ])

    # Row 5: Total Repayment Amount | (right box continues)
    repay_label_para = (
        _para([_run('Total Repayment Amount', bold=True, size=20)], align='left', spacing_after=20) +
        _para([_run('[Disbursement Amount plus (+) Finance Charge]', italic=True, size=15)],
              align='left', spacing_after=0)
    )
    row5 = _row([
        _cell(repay_label_para, LEFT_LBL_W, valign='top'),
        value_cell(pa_fmt, LEFT_VAL_W),
        right_cell_continue(),
    ])

    # Row 6: Estimated Time Period | (right box continues — last row of merge)
    row6 = _row([
        label_cell('Estimated Time Period Required for the Periodic Payments to Equal the Total Repayment Amount', LEFT_LBL_W),
        value_cell(str(term_count), LEFT_VAL_W),
        right_cell_continue(),
    ])

    # ── Payment Schedule block — spans full width, 4-col grid ──────────────
    # Reference shows a two-column-style block with checkboxes. We render it
    # as a single full-width cell containing the checkboxes.
    BOX_ON = '\u2612'   # ☒
    BOX_OFF = '\u25fb'  # ◻

    # Determine which checkbox to mark
    # The CT form has BOTH "Fixed Payment Amounts" and "Variable Payment Amounts"
    # sections. For MCAs, variable is checked and the method-description box is marked.
    pay_schedule_runs = (
        _run('Payment Schedule', bold=True, size=20) + '<w:br/>'.join([''])
    )
    # Build paragraphs inside the payment schedule cell
    ps_paras = []
    ps_paras.append(_para([_run('Payment Schedule', bold=True, size=20)], align='left', spacing_after=20))
    ps_paras.append(_para([_run('For Fixed Payment Amounts:', bold=True, size=20)], align='left', spacing_after=20))
    ps_paras.append(_para([_run(f'{BOX_OFF}  Amount of each estimated fixed payment: $', size=20)], align='left', spacing_after=20))
    ps_paras.append(_para([_run(f'{BOX_OFF}  Frequency of fixed payments: ', size=20)
                          + _run(freq_label_pg1, bold=True, size=20)], align='left', spacing_after=40))
    ps_paras.append(_para([_run('For Variable Payment Amounts:', bold=True, size=20)], align='left', spacing_after=20))
    ps_paras.append(_para([_run(f'{BOX_OFF}  Variable payment schedule, or $', size=20)], align='left', spacing_after=20))
    ps_paras.append(_para([_run(f'{BOX_ON}  Description of the method used to calculate the amount and frequency of each variable payment:', size=20)
                          + _run('   ', size=20)
                          + _run(f'{BOX_ON}  SEE PAGE 2', bold=True, size=20)], align='left', spacing_after=0))
    ps_cell = _cell(''.join(ps_paras), 10800, gridspan=4, valign='top')
    row_payments = _row([ps_cell])

    # ── Description of All Other Fees row ──────────────────────────────────
    fees_desc_left = _para(
        _run('Description of All Other Potential Fees and Charges ', bold=True, size=20)
        + _run('NOT', bold=True, underline=True, size=20)
        + _run(' Included in the Finance Charge (including draw fees, late payment fees, and returned payment fees)', bold=True, size=20),
        align='left', spacing_after=0)
    fees_desc_right = _para([_run(f'{BOX_ON}  SEE PAGE 2', bold=True, size=20)],
                            align='left', spacing_after=0)
    row_fees_desc = _row([
        _cell(fees_desc_left, LEFT_LBL_W + LEFT_VAL_W, gridspan=2, valign='center'),
        _cell(fees_desc_right, RIGHT_LBL_W + RIGHT_VAL_W, gridspan=2, valign='center'),
    ])

    # ── Description of Collateral row ──────────────────────────────────────
    collateral_label = _para([_run('Description of Collateral Requirements or Security Interests', bold=True, size=20)],
                             align='left', spacing_after=0)
    collateral_text = _para(_runs_rich(
        f'Provider will have a security interest in accounts and payment intangibles sold to '
        f'Provider by Recipient. This includes payments made by cash, check, ACH, or other '
        f'electronic transfer, credit card, debit card, bank card, charge card or other form '
        f'of monetary payment in the ordinary course of the Recipient\u2019s business, accounts '
        f'and payment intangibles and all proceeds of the foregoing.',
        size=20), align='left', spacing_after=20) + _para(
            [_run(f'{BOX_OFF}  SEE PAGE 2', size=20)], align='right', spacing_after=0)
    row_collateral = _row([
        _cell(collateral_label, LEFT_LBL_W, valign='top'),
        _cell(collateral_text, LEFT_VAL_W + RIGHT_LBL_W + RIGHT_VAL_W, gridspan=3, valign='top'),
    ])

    # ── Broker Compensation row ────────────────────────────────────────────
    broker_label = (
        _para([_run('Broker Compensation', bold=True, size=20)], align='left', spacing_after=20) +
        _para([_run('(Paid from Financed Amount)', size=16, italic=True)], align='left', spacing_after=0)
    )
    broker_yesno = (
        _para([_run('Is provider paying compensation directly to a broker?', size=20)],
              align='left', spacing_after=40) +
        _para([_run(f'{BOX_ON if broker_pays else BOX_OFF}  Yes   {BOX_OFF if broker_pays else BOX_ON}  No', size=20)],
              align='left', spacing_after=0)
    )
    broker_amt = (
        _para([_run('If yes, amount of compensation being paid directly to broker:', size=20)],
              align='left', spacing_after=40) +
        _para([_run(broker_amount_fmt, bold=True, underline=True, size=20)], align='left', spacing_after=0)
    )
    row_broker = _row([
        _cell(broker_label, LEFT_LBL_W, valign='top'),
        _cell(broker_yesno, LEFT_VAL_W + RIGHT_LBL_W // 2, gridspan=2, valign='top'),
        _cell(broker_amt, RIGHT_VAL_W + RIGHT_LBL_W // 2, valign='top'),
    ])

    # ── Prepayment row ────────────────────────────────────────────────────
    prepay_label = _para([_run(
        'Finance Charges or Fees upon Prepayment or Refinance ', bold=True, size=20)
        + _run('(including the percentage of any unpaid portion of the finance charge and the maximum dollar amount of finance charge)',
               italic=True, size=16)],
        align='left', spacing_after=0)
    prepay_text = _para(
        _run('If you pay off the financing early, you will still need to pay all or a portion of the '
             'finance charge, up to ', size=20)
        + _run(fc_fmt, bold=True, underline=True, size=20)
        + _run('. If you pay off the financing faster than required, you will not be required to '
               'pay additional fees.', size=20),
        align='left', spacing_after=20) + _para(
            [_run(f'{BOX_OFF}  SEE PAGE 2', size=20)], align='right', spacing_after=0)
    row_prepay = _row([
        _cell(prepay_label, LEFT_LBL_W, valign='top'),
        _cell(prepay_text, LEFT_VAL_W + RIGHT_LBL_W + RIGHT_VAL_W, gridspan=3, valign='top'),
    ])

    page1_table = _table(
        [row1, row2, row3, row4, row5, row6, row_payments, row_fees_desc,
         row_collateral, row_broker, row_prepay],
        total_width_dxa=10800,
    )

    # ── Signature row (outside the table, borderless 3-column grid) ───────
    # Three side-by-side: Merchant's Initials | Signature | Date.
    # Use a borderless table for clean alignment matching the reference.
    SIG_COL_INIT = 2200
    SIG_COL_SIG  = 5800
    SIG_COL_DATE = 2800

    def _sig_cell(line_chars, label, width):
        para = (
            _para([_run('_' * line_chars, size=20)], align='left', spacing_after=20) +
            _para([_run(label, size=16)], align='left', spacing_after=0)
        )
        cell_xml = '<w:tc>'
        cell_xml += '<w:tcPr>'
        cell_xml += f'<w:tcW w:w="{width}" w:type="dxa"/>'
        cell_xml += ('<w:tcBorders>'
                     '<w:top w:val="nil"/><w:left w:val="nil"/>'
                     '<w:bottom w:val="nil"/><w:right w:val="nil"/>'
                     '</w:tcBorders>')
        cell_xml += '<w:vAlign w:val="top"/>'
        cell_xml += '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        cell_xml += '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tcMar>'
        cell_xml += '</w:tcPr>'
        cell_xml += para
        cell_xml += '</w:tc>'
        return cell_xml

    # Tuned underscore counts so each line fits inside its column WITHOUT wrapping.
    # Initials col 2200 dxa ≈ 1.5", Signature col 5800 dxa ≈ 4", Date col 2800 dxa ≈ 1.9"
    sig_row_inner = (
        '<w:tr>'
        '<w:trPr><w:cantSplit/></w:trPr>'
        + _sig_cell(18, "Merchant's Initials", SIG_COL_INIT)
        + _sig_cell(48, 'Signature', SIG_COL_SIG)
        + _sig_cell(22, 'Date', SIG_COL_DATE)
        + '</w:tr>'
    )
    sig_table = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="10800" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblBorders>'
        '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        + sig_row_inner + '</w:tbl>'
    )
    # Empty paragraph above the signature table to push it down from the box.
    sig_spacer = _para([_run('', size=20)], spacing_before=120, spacing_after=0)
    sig_para = sig_spacer + sig_table
    sig_labels = ''  # labels are inside sig_table cells now

    # ── Page 2 — Title (must appear at bottom of page 1 per reference) ────
    # In the reference, "PAGE 2" title appears at the bottom of page 1 and
    # page 2 begins with the Recipient/Disclosure Date table. We'll put a
    # page break before page 2 content for cleanliness.
    pg2_title1 = _para([_run('COMMERCIAL (SALES-BASED) FINANCING DISCLOSURE FORM - PAGE 2',
                             bold=True, size=24)],
                       align='center', spacing_before=0, spacing_after=40)
    pg2_title2 = _para([_run('(pursuant to Part XVI of Chapter 669 of the 2024 Supplement to the General Statutes)',
                             size=20)],
                       align='center', spacing_after=120)

    # ── PAGE 2 — Header table inside the big box ──────────────────────────
    # In reference: page 2 is ONE bordered box containing the recipient/provider
    # header row, then all the body content.
    pg2_row1 = _row([
        kv_pair_cell("Recipient's Name:", merchant_name, 5400, gridspan=1),
        kv_pair_cell("Disclosure Date:", disclosure_date, 5400, gridspan=1),
    ])
    pg2_row2 = _row([
        kv_pair_cell("Recipient's Address:", merchant_addr, 5400, gridspan=1),
        kv_pair_cell("Provider's Name:", PROVIDER_NAME, 5400, gridspan=1),
    ])

    # All page-2 body content lives inside a single full-width cell that spans
    # both columns (gridspan=2). The outer table provides the bordered box.
    BOX_W = 10800

    # Checked items
    pg2_body = []
    pg2_body.append(_para([_run('The information provided below relates to the following checked item(s):',
                                bold=True, size=20)], align='left', spacing_before=60, spacing_after=40))
    pg2_body.append(_para([_run(f'{BOX_OFF}  Variable payment schedule', size=20)], align='left', indent_left=300, spacing_after=20))
    pg2_body.append(_para([_run(f'{BOX_ON}  Description of the method used to calculate the amount and frequency of each variable payment', size=20)], align='left', indent_left=300, spacing_after=20))
    pg2_body.append(_para([_run(f'{BOX_ON}  Description of all other potential fees and charges not included in the finance charge', size=20)], align='left', indent_left=300, spacing_after=20))
    pg2_body.append(_para([_run(f'{BOX_OFF}  Description of collateral requirements or security interests', size=20)], align='left', indent_left=300, spacing_after=20))
    pg2_body.append(_para([_run(f'{BOX_OFF}  Description of prepayment policies, finance charges and fees', size=20)], align='left', indent_left=300, spacing_after=120))

    # Method heading & body — frequency auto-selected
    pg2_body.append(_para([_run('Description of the method used to calculate the amount and frequency of each variable payment',
                                bold=True, underline=True, size=20)],
                         align='left', spacing_after=20))
    pg2_body.append(_para(_runs_rich(
        f'Provider will collect the Total Repayment Amount by debiting Recipient\u2019s business '
        f'bank account ("payments") that will occur with the following frequency (the option '
        f'marked {BOX_ON} applies):',
        size=20), align='left', spacing_after=40))

    daily_checked = (not is_weekly)
    weekly_checked = is_weekly
    pg2_body.append(_para([_run(f'{BOX_ON if daily_checked else BOX_OFF}  Every Business Day (i.e., Monday through Friday, excluding bank holidays)', size=20)], align='left', indent_left=300, spacing_after=20))
    pg2_body.append(_para([_run(f'{BOX_ON if weekly_checked else BOX_OFF}  Weekly', size=20)], align='left', indent_left=300, spacing_after=20))
    pg2_body.append(_para([_run(f'{BOX_OFF}  Monthly', size=20)], align='left', indent_left=300, spacing_after=80))

    # Initial payment / % / monthly — UNDERLINED per Jack's spec
    pg2_body.append(_para(
        _run('The initial payment will be ', size=20)
        + _run(pmt_fmt, bold=True, underline=True, size=20)
        + _run('. We based your initial payment on ', size=20)
        + _run(f'{spec_pct:.2f}%', bold=True, underline=True, size=20)
        + _run(' of your estimated sales revenue. For details on your right to adjust any payment '
               'amount, see Page 3 of your Purchase Agreement. Based on this information, the amount '
               'of the average projected payments per month is ', size=20)
        + _run(avg_monthly_fmt, bold=True, underline=True, size=20)
        + _run('.', size=20),
        align='left', spacing_after=120))

    # Fees heading & body — returned entry fee $50 UNDERLINED
    pg2_body.append(_para([_run('Description of all other potential fees and charges not included in the finance charge',
                                bold=True, underline=True, size=20)],
                         align='left', spacing_after=20))
    pg2_body.append(_para(
        _run('If any ACH entry, check, or electronically created item is returned or rejected for '
             'insufficient funds, then Recipient will pay Provider a returned entry fee of ', size=20)
        + _run('$50', bold=True, underline=True, size=20)
        + _run('. In addition, Recipient will reimburse Provider for any charges incurred by Provider '
               'resulting from the returned ACH entry, check, or electronically created item. Include '
               'charges for breach, financing statement filing fee, or any default charges if applicable].',
               size=20),
        align='left', spacing_after=120))

    # Renewal section — pull table inline as paragraphs to keep inside the box
    pg2_body.append(_para([_run('If a renewal financing transaction:', bold=True, underline=True, size=20)],
                         align='left', spacing_after=40))

    # Inline renewal mini-table (nested table inside the body cell)
    def _renewal_inline_row(label, value):
        return _row([
            _cell(_para([_run(label, size=20)], align='left', spacing_after=0), 7800, valign='center'),
            _cell(_para([_run(value, size=20)], align='left', spacing_after=0), 2800, valign='center'),
        ])

    nested_renewal_tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="10600" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        + _renewal_inline_row('Reduction in Disbursement Amount to Pay Outstanding Balance of Existing Commercial Financing *', '$0.00')
        + _renewal_inline_row('Total Amount New Financing Used to Payoff Prepayment Charges & Unpaid Interest on Existing Commercial Financing', '$0.00')
        + _renewal_inline_row('   \u2022 Prepayment Charges Payable to Provider', '$0.00')
        + _renewal_inline_row('   \u2022 Unpaid Interest Payable to Provider Not Forgiven at the time of Renewal', '$0.00')
        + '</w:tbl>'
    )
    pg2_body.append(nested_renewal_tbl)
    pg2_body.append(_para([_run('', size=12)], spacing_after=40))  # small spacer

    # Statutory boilerplate
    pg2_body.append(_para([_run(
        'Connecticut law prohibits commercial financing contracts from having any provision '
        'waiving a Recipient\u2019s right to notice, judicial hearing, or prior court order '
        'under Chapter 903a in connection with the Provider obtaining a prejudgment remedy, '
        'upon commencing any litigation against the Recipient.',
        size=20)], align='left', spacing_after=40))
    pg2_body.append(_para([_run(
        'Provider will not revoke, withdraw, or modify a specific offer for commercial '
        'financing made until midnight of the third calendar day after the date of this '
        'offer. A specific offer may be revoked, withdrawn or modified: (1) based on '
        'information obtained in the underwriting process, including but not limited to, '
        'verification of any information provided by the Recipient, or (2) at the request '
        'of the Recipient.',
        size=20)], align='left', spacing_after=40))
    pg2_body.append(_para([_run(
        'This specific offer for commercial financing is (1) based on the provider\u2019s '
        'preliminary review of application information only and (2) not a final approval '
        'or commitment to provide commercial financing.',
        size=20)], align='left', spacing_after=80))

    # Acknowledgement and signature
    pg2_body.append(_para(
        _run('Acknowledgement: ', bold=True, size=20)
        + _run('I/We acknowledge that I/we have received this Commercial Financing Disclosure Form.', size=20),
        align='left', spacing_before=40, spacing_after=120))
    pg2_body.append(_para(
        _run('Signature of Merchant: ', bold=True, size=20)
        + _run('_' * 50, size=20),
        align='left', spacing_after=120))
    if two_signers and signer2_name:
        pg2_body.append(_para(
            _run('Signature of Merchant: ', bold=True, size=20)
            + _run('_' * 50, size=20),
            align='left', spacing_after=120))
    pg2_body.append(_para(
        _run('Date: ', bold=True, size=20)
        + _run('_' * 25, size=20),
        align='left', spacing_after=40))

    # Body row — spans 2 columns
    pg2_body_row = _row([
        _cell(''.join(pg2_body), BOX_W, gridspan=2, valign='top'),
    ])

    # ONE single bordered box for ALL of page 2
    pg2_box = _table([pg2_row1, pg2_row2, pg2_body_row], total_width_dxa=BOX_W)

    # Stub legacy variables so the body assembly still compiles
    pg2_header_table = ''
    checked_items = ''
    method_heading = ''
    method_body1 = ''
    freq_lines = ''
    method_body2 = ''
    fees_heading = ''
    fees_body = ''
    renewal_heading = ''
    renewal_table = ''
    boilerplate = ''
    ack = ''
    sig2 = ''
    date2 = ''

    # ── Section properties (page margins ~0.75") ──────────────────────────
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )

    # ── Assemble body ──────────────────────────────────────────────────────
    body = (
        title1 + title2
        + page1_table
        + sig_para + sig_labels
        + _page_break()
        + pg2_title1 + pg2_title2
        + pg2_box
        + sect_pr
    )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>\n'
        f'<w:body>{body}</w:body>\n'
        '</w:document>'
    )

    # ── DOCX scaffolding ──────────────────────────────────────────────────
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
