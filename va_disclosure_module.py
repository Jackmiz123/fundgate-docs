"""
VA (Virginia) Sales-Based Financing Disclosure module — builds the FundGate LLC
Virginia disclosure as DOCX bytes.

Format mirrors the executed KOVANES CPA reference disclosure exactly:
Page 1 — Single bordered table (left: financial details, right: merged
         provider/recipient box), payment-schedule checkboxes + narrative,
         other-fees / collateral / broker-compensation / prepayment rows,
         then acknowledgement + signature line.
Page 2 — Header box (recipient/provider) + body box: checked items, how the
         increment is determined, prepayment language, additional-fees note,
         outstanding-balance note, acknowledgement + signature line.

FundGate-only. Provider name/address/phone/email are hardcoded.
"""
import io, zipfile
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# Value helpers
# ─────────────────────────────────────────────────────────────────────────────
def _fmt_currency(val):
    try:
        n = float(str(val).replace('$', '').replace(',', '').replace('%', ''))
        return f'${n:,.2f}'
    except Exception:
        return str(val) if val else '$0.00'


def _fmt_date_slash(val):
    """Virginia reference uses M/D/YYYY (e.g. 6/15/2026)."""
    if not val:
        return ''
    for fmt in ('%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d'):
        try:
            dt = datetime.strptime(str(val).strip(), fmt)
            return f'{dt.month}/{dt.day}/{dt.year}'
        except Exception:
            pass
    return str(val)


def _n(data, key, default=0.0):
    try:
        return float(str(data.get(key, default)).replace('$', '').replace(',', '').replace('%', ''))
    except Exception:
        return default


def _first(data, keys, default=''):
    for k in keys:
        v = data.get(k)
        if v not in (None, ''):
            return v
    return default


# ─────────────────────────────────────────────────────────────────────────────
# XML helpers (shared house style: Times New Roman, thin black borders)
# ─────────────────────────────────────────────────────────────────────────────
FONT = 'Times New Roman'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
BOX_ON = '☒'   # ☒
BOX_OFF = '☐'  # ☐


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
    tcpr += '<w:tcMar><w:top w:w="140" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
    tcpr += '<w:bottom w:w="140" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'
    tcpr += '</w:tcPr>'
    return f'<w:tc>{tcpr}{content_xml}</w:tc>'


def _cellnb(content_xml, width_dxa, gridspan=None, valign='top'):
    """Borderless cell (for signature blocks)."""
    tcpr = '<w:tcPr>'
    tcpr += f'<w:tcW w:w="{width_dxa}" w:type="dxa"/>'
    if gridspan:
        tcpr += f'<w:gridSpan w:val="{gridspan}"/>'
    tcpr += ('<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>'
             '<w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>')
    if valign:
        tcpr += f'<w:vAlign w:val="{valign}"/>'
    tcpr += ('<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
             '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>')
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
# Provider constants — FundGate LLC (hardcoded)
# ─────────────────────────────────────────────────────────────────────────────
PROVIDER_NAME  = 'FundGate LLC'
PROVIDER_ADDR  = '1202 Avenue U, Suite 1175, Brooklyn, NY 11229'
PROVIDER_PHONE = '631-772-9020'
PROVIDER_EMAIL = 'admin@fundgatellc.com'


def build_va_disclosure_bytes(data):
    """Build the FundGate Virginia sales-based financing disclosure as DOCX bytes."""
    # ── Recipient ─────────────────────────────────────────────────────────
    _legal = (data.get('Merchant_Legal_Name', '') or '').strip()
    _dba   = (data.get('Merchant_DBA', '') or '').strip()
    if _dba and _dba.upper() != _legal.upper():
        merchant_name = f'{_legal} DBA {_dba}'.upper()
    else:
        merchant_name = _legal.upper()
    merchant_addr = (data.get('Executive_Office_Address', '') or '').strip()
    disclosure_date = _fmt_date_slash(data.get('Agreement_Date', ''))

    # ── Amounts ───────────────────────────────────────────────────────────
    pp = _n(data, 'Purchase_Price')      # Total Amount of the Sales-Based Financing
    pa = _n(data, 'Purchased_Amount')    # Total Repayment Amount

    # Fees withheld at disbursement (supports % or $ modes, like the other states)
    ach_pct = _n(data, 'ACH_Program_Fee_Percentage')
    orig_pct = _n(data, 'Origination_Fee_Percentage')
    ach_fee_mode = (data.get('ACH_Program_Fee_Mode', 'pct') or 'pct').lower()
    orig_fee_mode = (data.get('Origination_Fee_Mode', 'pct') or 'pct').lower()
    ach_fee = ach_pct if ach_fee_mode == 'dollar' else round(pp * ach_pct / 100, 2)
    orig_fee = orig_pct if orig_fee_mode == 'dollar' else round(pp * orig_pct / 100, 2)
    total_fees = _n(data, 'Fees_Withheld') if data.get('Fees_Withheld') else round(ach_fee + orig_fee, 2)

    prior = _n(data, 'Prior_Balance_Amount')             # balances paid off at funding

    # Disbursement Amount = amount actually paid to the merchant (net of fees AND prior balances)
    disbursement = round(pp - total_fees - prior, 2)
    # Finance Charge = Total Repayment − (Purchase Price − Fees withheld)
    finance_charge = round(pa - (pp - total_fees), 2)

    # ── Payment schedule ──────────────────────────────────────────────────
    ach_freq = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()
    is_weekly = 'week' in ach_freq
    incr_word = 'Weekly' if is_weekly else 'Daily'
    if is_weekly:
        pmt = _n(data, 'Specific_Weekly_Amount')
    else:
        pmt = _n(data, 'Specific_Daily_Amount')

    try:
        num_payments = int(float(str(data.get('Payment_Term_Count', '') or '').replace(',', '').strip()))
    except Exception:
        num_payments = int(round(pa / pmt)) if (pmt > 0 and pa > 0) else 0

    # ── Broker compensation (accept VA-, CT-, or generic-named form fields) ─
    broker_pays = str(_first(data, ['VA_Broker_Pays', 'CT_Broker_Pays', 'Broker_Pays'], '')).strip().lower() in ('yes', 'true', '1', 'y')
    broker_amount = _n(data, 'VA_Broker_Amount') or _n(data, 'CT_Broker_Amount') or _n(data, 'Broker_Amount')

    # ── Formatted strings ─────────────────────────────────────────────────
    pp_fmt = _fmt_currency(pp)
    pa_fmt = _fmt_currency(pa)
    fees_fmt = _fmt_currency(total_fees)
    disb_fmt = _fmt_currency(disbursement)
    fc_fmt = _fmt_currency(finance_charge)
    pmt_fmt = _fmt_currency(pmt)
    prior_fmt = _fmt_currency(prior)
    broker_amount_fmt = _fmt_currency(broker_amount) if broker_pays else '$0.00'

    # ── Column geometry (total 10800 dxa) ─────────────────────────────────
    C1, C2, C3, C4 = 2900, 2100, 2300, 3500
    RIGHT_W = C3 + C4

    def L(text, w, gridspan=None, valign='center'):
        return _cell(_para([_run(text, bold=True, size=20)], align='left', spacing_after=0),
                     w, gridspan=gridspan, valign=valign)

    def V(text, w, gridspan=None):
        return _cell(_para([_run(text, bold=True, size=20)], align='left', spacing_after=0),
                     w, gridspan=gridspan, valign='center')

    # Right merged box content
    def kv(label, value, after=220):
        return _para(_run(label, bold=True, size=20) + _run(' ' + value, size=20),
                     align='left', spacing_after=after)

    right_box = (
        kv('Disclosure Date:', disclosure_date) +
        kv("Recipient's Name:", merchant_name) +
        kv("Recipient's Address:", merchant_addr) +
        kv("Provider's Name:", PROVIDER_NAME) +
        kv("Provider's Address:", PROVIDER_ADDR) +
        kv("Provider's Phone Number:", PROVIDER_PHONE) +
        kv("Provider's E-mail Address:", PROVIDER_EMAIL, after=0)
    )

    def rbox_restart():
        return _cell(right_box, RIGHT_W, gridspan=2, valign='top', vmerge='restart')

    def rbox_continue():
        return _cell(_para([_run('', size=20)]), RIGHT_W, gridspan=2, valign='top', vmerge='continue')

    RH = 620  # min row height so the left rows match the taller reference layout
    row1 = _row([L('Total Amount of the Sales-Based Financing', C1), V(pp_fmt, C2), rbox_restart()], height_dxa=RH)
    row2 = _row([L('Fees Deducted or Withheld at Disbursement', C1), V(fees_fmt, C2), rbox_continue()], height_dxa=RH)
    row3 = _row([L('Disbursement Amount', C1), V(disb_fmt, C2), rbox_continue()], height_dxa=RH)
    row4 = _row([L('Finance Charge', C1), V(fc_fmt, C2), rbox_continue()], height_dxa=RH)
    row5 = _row([L('Total Repayment Amount', C1), V(pa_fmt, C2), rbox_continue()], height_dxa=RH)
    row6 = _row([L('Estimated Number of Payments', C1), V(str(num_payments), C2), rbox_continue()], height_dxa=RH)

    # ── Payment Schedule row ──────────────────────────────────────────────
    ps_left = (
        _para([_run('Payment Schedule', bold=True, size=20)], align='left', spacing_after=160) +
        _para([_run(f'{BOX_ON} ', size=20) + _run('Fixed', bold=True, size=20)], align='left', spacing_after=0)
    )
    freq_narrative = (
        _run('Payment frequency', bold=True, size=20)
        + _run(': If Daily: Provider will debit the Daily Increment from Recipient’s business bank '
               'account each business day [Monday-Friday]. The initial Daily Increment is ____________ '
               'If the date of any payments falls on a bank holiday that payment will be made up the next '
               'business day in addition to the regularly scheduled payment for that day '
               'If Weekly: Provider will debit the Weekly Increment from Recipient’s business bank '
               'account on the designated business day each week The initial Weekly Increment is ', size=20)
        + _run(pmt_fmt, bold=True, size=20)
        + _run('. Please refer to Page 4 of your Agreement for how the Weekly Increment can be changed. '
               'If the date of any payments falls on a bank holiday that payment will be made up the next '
               'business day ', size=20)
        + _run('Method of payment', bold=True, size=20)
        + _run(': ACH debit from Recipient’s business bank account.', size=20)
    )
    ps_right = (
        _para([_run(f'{BOX_ON}  Amount of each estimated fixed payment: ', size=20) + _run(pmt_fmt, size=20)],
              align='left', spacing_after=40) +
        _para([_run(f'{BOX_OFF}  Frequency of fixed payments: ', size=20) + _run(incr_word, bold=True, size=20)],
              align='left', spacing_after=120) +
        _para(freq_narrative, align='left', spacing_after=60) +
        _para([_run(f'{BOX_ON}  SEE PAGE 2', bold=True, size=20)], align='right', spacing_after=0)
    )
    row_pay = _row([
        _cell(ps_left, C1, valign='top'),
        _cell(ps_right, C2 + C3 + C4, gridspan=3, valign='top'),
    ])

    # ── Other fees row ────────────────────────────────────────────────────
    fees_left = _para(
        _run('Description of All Other Potential Fees and Charges ', bold=True, size=20)
        + _run('NOT', bold=True, underline=True, size=20)
        + _run(' Included in the Finance Charge', bold=True, size=20),
        align='left', spacing_after=0)
    row_fees = _row([
        _cell(fees_left, C1 + C2, gridspan=2, valign='center'),
        _cell(_para([_run(f'{BOX_ON}  SEE PAGE 2', bold=True, size=20)], align='left'),
              C3 + C4, gridspan=2, valign='center'),
    ])

    # ── Collateral row ────────────────────────────────────────────────────
    collat_right = (
        _para([_run('Provider has a security interest in certain accounts and payment intangibles sold to '
                    'Provider by Recipient.', size=20)], align='left', spacing_after=40) +
        _para([_run(f'{BOX_OFF}  SEE PAGE 2', bold=True, size=20)], align='right', spacing_after=0)
    )
    row_collat = _row([
        L('Description of Collateral Requirements or Security Interests', C1 + C2, gridspan=2, valign='top'),
        _cell(collat_right, C3 + C4, gridspan=2, valign='top'),
    ])

    # ── Broker compensation row ───────────────────────────────────────────
    broker_mid = (
        _para([_run('Is provider paying compensation directly to a broker?', size=20)],
              align='left', spacing_after=120) +
        _para([_run(f'{BOX_ON if broker_pays else BOX_OFF}  Yes    {BOX_OFF if broker_pays else BOX_ON}  No', size=20)],
              align='left', spacing_after=0)
    )
    broker_right = (
        _para([_run('If yes, amount of compensation being paid directly to broker:', size=20)],
              align='left', spacing_after=120) +
        _para([_run(broker_amount_fmt, bold=True, underline=True, size=20)], align='left', spacing_after=0)
    )
    row_broker = _row([
        L('Broker Compensation', C1, valign='top'),
        _cell(broker_mid, C2 + C3, gridspan=2, valign='top'),
        _cell(broker_right, C4, valign='top'),
    ])

    # ── Prepayment row ────────────────────────────────────────────────────
    row_prepay = _row([
        L('Description of Prepayment Policies', C1 + C2, gridspan=2, valign='center'),
        _cell(_para([_run(f'{BOX_ON}  SEE PAGE 2', bold=True, size=20)], align='left'),
              C3 + C4, gridspan=2, valign='center'),
    ])

    page1_table = _table(
        [row1, row2, row3, row4, row5, row6, row_pay, row_fees, row_collat, row_broker, row_prepay],
        total_width_dxa=10800)

    # ── Acknowledgement + signature (page 1, borderless) ──────────────────
    def sig_block():
        ack = _para([_run('I acknowledge that I have received a copy of this disclosure form.', size=20)],
                    align='left', spacing_before=160, spacing_after=200)
        left_cell = _cellnb(
            _para([_run('', size=20)], align='left', spacing_after=60) +
            _para([_run('_' * 40, size=20)], align='left', spacing_after=20) +
            _para([_run('Signature', size=20)], align='left', spacing_after=0),
            5400, valign='top')
        right_cell = _cellnb(
            _para([_run(disclosure_date, bold=True, size=20)], align='left', spacing_after=60) +
            _para([_run('_' * 28, size=20)], align='left', spacing_after=20) +
            _para([_run('Date', size=20)], align='left', spacing_after=0),
            5400, valign='top')
        sig_tbl = (
            '<w:tbl><w:tblPr><w:tblW w:w="10800" w:type="dxa"/><w:jc w:val="center"/>'
            '<w:tblLayout w:type="fixed"/>'
            '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
            '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
            '</w:tblPr>' + _row([left_cell, right_cell]) + '</w:tbl>'
        )
        return ack + sig_tbl

    sig1 = sig_block()

    # ── Page 2 ────────────────────────────────────────────────────────────
    title1 = _para([_run('SALES-BASED FINANCING DISCLOSURE FORM', bold=True, size=26)],
                   align='center', spacing_before=0, spacing_after=160)
    title2 = _para([_run('SALES-BASED FINANCING DISCLOSURE FORM - PAGE 2', bold=True, size=26)],
                   align='center', spacing_before=0, spacing_after=160)

    def kvcell(label, value, w):
        return _cell(_para(_run(label, bold=True, size=20) + _run(' ' + value, size=20),
                           align='left', spacing_after=0), w, valign='top')

    pg2_header = _table([
        _row([kvcell("Recipient's Name:", merchant_name, 5400), kvcell("Disclosure Date:", disclosure_date, 5400)]),
        _row([kvcell("Recipient's Address:", merchant_addr, 5400), kvcell("Provider's Name:", PROVIDER_NAME, 5400)]),
    ], total_width_dxa=10800)

    body = []
    body.append(_para([_run('The information provided below relates to the following checked item(s):', bold=True, size=20)],
                      align='left', spacing_before=40, spacing_after=240))
    body.append(_para([_run(f'{BOX_OFF}  Variable payment schedule', size=20)], align='left', spacing_after=60))
    body.append(_para([_run(f'{BOX_ON}  Description of the method used to calculate the amount and frequency of each fixed payment', size=20)],
                      align='left', spacing_after=220))
    body.append(_para([_run(f'{BOX_OFF}  Method of payment', size=20)], align='left', spacing_after=60))
    body.append(_para(_run(f'{BOX_ON}  Description of all other potential fees and charges ', size=20)
                      + _run('not', underline=True, size=20)
                      + _run(' included in the finance charge.', size=20),
                      align='left', spacing_after=220))
    body.append(_para([_run(f'{BOX_OFF}  Description of collateral requirements or security interests', size=20)],
                      align='left', spacing_after=60))
    body.append(_para([_run(f'{BOX_ON}  Description of prepayment policies', size=20)], align='left', spacing_after=260))

    body.append(_para(
        _run(f'How we determine the {incr_word} Increment', bold=True, size=20)
        + _run(': We review information you provide or make available to us including but not limited to the '
               'bank statements showing the revenue into your business bank account to calculate your total '
               'sales revenue over a period of time. We then estimate the average amount of your sales revenue '
               'per the payment frequency described above. Then we multiply this amount by the Purchase '
               f'Percentage described in your Agreement to determine your initial {incr_word} Increment.', size=20),
        align='left', spacing_after=200))

    body.append(_para([_run('If you pay off the financing faster than required, you will not be required to pay additional fees', size=20)],
                      align='left', spacing_after=80))
    body.append(_para([_run('If you pay off the financing faster than required, and there is a prepayment affidavit entered at '
                            'funding you will receive a discount for prepayment. If there is no discounted prepayment the Total '
                            'Repayment Amount must be paid', size=20)],
                      align='left', spacing_after=200))

    body.append(_para([_run('ADDITIONAL FEES WHICH MAY APPLY: See Schedule to the Contract', bold=True, size=20)],
                      align='left', indent_left=360, spacing_after=200))

    body.append(_para([_run('* Which includes any balances currently outstanding being paid to FundGate, LLC, and/or third '
                            'parties, in the sum of: ' + prior_fmt, size=20)],
                      align='left', spacing_after=80))
    body.append(_para([_run('I acknowledge that I have received a copy of this disclosure form.', size=20)],
                      align='left', spacing_after=260))

    # page-2 signature line (borderless, matches reference)
    sig2_left = _cellnb(
        _para([_run('', size=20)], align='left', spacing_after=60) +
        _para([_run('_' * 40, size=20)], align='left', spacing_after=20) +
        _para([_run('Signature', size=20)], align='left', spacing_after=0),
        5400, valign='top')
    sig2_right = _cellnb(
        _para([_run(disclosure_date, bold=True, size=20)], align='left', spacing_after=60) +
        _para([_run('_' * 28, size=20)], align='left', spacing_after=20) +
        _para([_run('Date', size=20)], align='left', spacing_after=0),
        5400, valign='top')
    sig2_tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="10800" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
        '</w:tblPr>' + _row([sig2_left, sig2_right]) + '</w:tbl>'
    )
    body.append(_para([_run('', size=12)], spacing_after=160))
    body.append(sig2_tbl)
    body.append(_para([_run('', size=8)]))   # a table cell must END with a paragraph, not a nested table

    # One big bordered box wrapping the whole page-2 body (matches reference)
    pg2_big_box = _table([_row([_cell(''.join(body), 10800, valign='top')])], total_width_dxa=10800)

    # ── Section properties ────────────────────────────────────────────────
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )

    body_xml = (
        title1
        + page1_table
        + sig1
        + _page_break()
        + title2
        + pg2_header
        + _para([_run('', size=8)], spacing_after=120)
        + pg2_big_box
        + sect_pr
    )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>\n'
        f'<w:body>{body_xml}</w:body>\n'
        '</w:document>'
    )

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
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
              '</w:docDefaults></w:styles>')

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document_xml)
        z.writestr('word/styles.xml', styles)
    return buf.getvalue()
