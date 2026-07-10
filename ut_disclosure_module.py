"""
UT Commercial Financing Disclosure module — builds the FundGate LLC Utah
disclosure as DOCX bytes.

Required by the Utah Commercial Financing Registration and Disclosure Act,
Utah Code Title 7, Chapter 27 (Section 7-27-202).

Format mirrors the BizFund reference disclosure: a single page.

  Left block  — Total Amount of the Commercial Financing, Fees Deducted or
                Withheld at Disbursement, Total Amount of Funds Disbursed,
                Total Amount to be Paid to Us, Total Dollar Cost.
  Right block — statutory intro sentence, broker name and compensation,
                disclosure date, merchant and provider identity.
  Bottom      — Initial Estimated Payment Amount, Description of Prepayment
                Policies, acknowledgment, signature and date lines.

This module is FundGate-only by design — the UT disclosure is never branded
as Fundkey. Provider name, address, phone, and email are hardcoded.

Broker name and broker compensation come from the form inputs
UT_Broker_Name and UT_Broker_Amount. Utah requires disclosure of funds paid
to a broker regardless of which party bears the cost, so the broker amount
does NOT affect the disbursement math.
"""
import io
import zipfile
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
    """UT format uses MM-DD-YYYY with dashes (e.g. 07-10-2026)."""
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


def _run(text, bold=False, italic=False, size=21, underline=False):
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


def _runs_rich(text, size=21, italic=False):
    """Parse a string with **bold** inline markers into multiple <w:r> runs."""
    parts, current, is_bold, i = [], [], False, 0
    while i < len(text):
        if text[i:i + 2] == '**':
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
    return ''.join(_run(t, bold=b, size=size, italic=italic) for t, b in parts if t)


def _para(runs_xml, align=None, spacing_before=0, spacing_after=0, line=None):
    ppr = '<w:pPr>'
    if align:
        ppr += f'<w:jc w:val="{align}"/>'
    sp = ''
    if spacing_before or spacing_after:
        sp += f'w:before="{spacing_before}" w:after="{spacing_after}" '
    if line:
        sp += f'w:line="{line}" w:lineRule="auto" '
    if sp:
        ppr += f'<w:spacing {sp.strip()}/>'
    ppr += '</w:pPr>'
    if isinstance(runs_xml, (list, tuple)):
        runs_xml = ''.join(runs_xml)
    return f'<w:p>{ppr}{runs_xml}</w:p>'


def _empty_para(size=16):
    return f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>{_run("", size=size)}</w:p>'


def _cell(content_xml, width_dxa, gridspan=None, valign='center', vmerge=None):
    tcpr = '<w:tcPr>'
    tcpr += f'<w:tcW w:w="{width_dxa}" w:type="dxa"/>'
    if gridspan:
        tcpr += f'<w:gridSpan w:val="{gridspan}"/>'
    if vmerge:
        tcpr += (f'<w:vMerge w:val="{vmerge}"/>' if vmerge == 'restart'
                 else '<w:vMerge/>')
    tcpr += ('<w:tcBorders>'
             '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
             '</w:tcBorders>')
    if valign:
        tcpr += f'<w:vAlign w:val="{valign}"/>'
    tcpr += ('<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
             '<w:bottom w:w="100" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>')
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
# Provider constants — FundGate LLC only
# ─────────────────────────────────────────────────────────────────────────────
PROVIDER_NAME  = 'FundGate LLC'
PROVIDER_ADDR  = '1202 Avenue U, Suite 1175, Brooklyn, NY 11229'
PROVIDER_PHONE = '631-772-9020'
PROVIDER_EMAIL = 'admin@fundgatellc.com'

# Column widths (dxa, 1440 per inch) — total 10800 = 7.5"
COL_LABEL = 3860
COL_VALUE = 1900
COL_INFO  = 5040


def build_ut_disclosure_bytes(data):
    """
    Build the FundGate UT commercial financing disclosure as DOCX bytes.
    Returns None if state is not UT.
    """
    state_code = (data.get('State_of_Organization') or '').upper().strip()
    if state_code != 'UT':
        return None

    # ── Pull inputs ─────────────────────────────────────────────────────────
    _legal = (data.get('Merchant_Legal_Name', '') or '').strip()
    _dba = (data.get('Merchant_DBA', '') or '').strip()
    if _dba and _dba.upper() != _legal.upper():
        merchant_name = f'{_legal} DBA {_dba}'.upper()
    else:
        merchant_name = _legal.upper()
    merchant_addr = (data.get('Executive_Office_Address', '') or '').strip()
    disclosure_date = _fmt_date_dash(data.get('Agreement_Date', ''))

    pp = _n(data, 'Purchase_Price')
    pa = _n(data, 'Purchased_Amount')

    ach_pct = _n(data, 'ACH_Program_Fee_Percentage')
    orig_pct = _n(data, 'Origination_Fee_Percentage')
    ach_fee_mode = (data.get('ACH_Program_Fee_Mode', 'pct') or 'pct').lower()
    orig_fee_mode = (data.get('Origination_Fee_Mode', 'pct') or 'pct').lower()
    ach_fee = ach_pct if ach_fee_mode == 'dollar' else round(pp * ach_pct / 100, 2)
    orig_fee = orig_pct if orig_fee_mode == 'dollar' else round(pp * orig_pct / 100, 2)

    total_fees = round(ach_fee + orig_fee, 2)
    disbursement = round(pp - total_fees, 2)
    # Utah 7-27-202: total dollar cost = total to be paid minus amount provided.
    total_dollar_cost = round(pa - pp, 2)

    # Payment frequency
    ach_freq = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()
    is_weekly = 'week' in ach_freq
    pmt = _n(data, 'Specific_Weekly_Amount') if is_weekly else _n(data, 'Specific_Daily_Amount')
    per_label = 'per week' if is_weekly else 'per day'

    # Broker (Utah requires disclosure of funds paid to brokers)
    broker_name = (data.get('UT_Broker_Name', '') or '').strip() or 'None'
    broker_amount = _fmt_currency(_n(data, 'UT_Broker_Amount'))

    # Formatted strings
    pp_fmt = _fmt_currency(pp)
    fees_fmt = _fmt_currency(total_fees)
    disb_fmt = _fmt_currency(disbursement)
    pa_fmt = _fmt_currency(pa)
    tdc_fmt = _fmt_currency(total_dollar_cost)
    pmt_fmt = _fmt_currency(pmt)
    third_party_fmt = _fmt_currency(_n(data, 'Prior_Balance_Amount'))
    direct_fmt = _fmt_currency(round(disbursement - _n(data, 'Prior_Balance_Amount'), 2))

    # ── Title ───────────────────────────────────────────────────────────────
    title = _para(_run('COMMERCIAL FINANCING DISCLOSURE FORM', bold=True, size=25,
                       underline=True), align='center', spacing_after=360)

    # ── Right-hand info block ───────────────────────────────────────────────
    intro = (f'This Commercial Financing Disclosure Form is being provided by '
             f'**({PROVIDER_NAME})** (\u201cwe\u201d or \u201cus\u201d) to '
             f'**{merchant_name}** (\u201cyou\u201d) as required by Utah law.')
    LN = 360  # 1.5 line spacing, matches the reference
    info_paras = [
        _para(_runs_rich(intro), spacing_after=180, line=LN),
        _para(_runs_rich(f'Name of Broker: **{broker_name}**'), line=LN),
        _para(_runs_rich(f'Amount paid to Broker: **{broker_amount}**'), line=LN),
        _para(_runs_rich(f'Disclosure Date: **{disclosure_date}**'), line=LN),
        _para(_runs_rich(f'Merchant\u2019s Name: **{merchant_name}**'), line=LN),
        _para(_runs_rich(f'Merchant\u2019s Address: **{merchant_addr}**'), line=LN),
        _para(_runs_rich(f'Provider\u2019s Name: **{PROVIDER_NAME}**'), line=LN),
        _para(_runs_rich(f'Provider\u2019s Address: **{PROVIDER_ADDR}**'), line=LN),
        _para([_run('Provider\u2019s Phone Number: '),
               _run(PROVIDER_PHONE, bold=True, italic=True)], line=LN),
        _para(_run('Provider\u2019s E-mail Address:'), line=LN),
        _para(_run(PROVIDER_EMAIL, bold=True, italic=True), line=LN),
    ]
    info_cell = _cell(''.join(info_paras), COL_INFO, valign='top', vmerge='restart')
    info_cont = _cell(_empty_para(), COL_INFO, valign='top', vmerge='continue')

    # ── Left rows ───────────────────────────────────────────────────────────
    r1 = _row([
        _cell(_para(_run('Total Amount of the Commercial Financing', bold=True)), COL_LABEL),
        _cell(_para(_run(pp_fmt)), COL_VALUE),
        info_cell,
    ], height_dxa=1440)

    r2 = _row([
        _cell(_para(_run('Fees Deducted or Withheld at Disbursement', bold=True)), COL_LABEL),
        _cell(_para(_run(fees_fmt)), COL_VALUE),
        info_cont,
    ], height_dxa=1440)

    disb_cell_xml = (
        _para(_run('Total Amount of Funds Disbursed', bold=True))
        + _para(_run('Total Amount of the Commercial Financing - Fees Deducted or '
                     'Withheld at Disbursement', size=18))
        + _empty_para(size=12)
        + _para(_runs_rich('Amount paid on your account with us or paid on your behalf '
                           f'to third parties **{third_party_fmt}**', size=18))
        + _empty_para(size=12)
        + _para(_runs_rich(f'Amount paid directly to you **{direct_fmt}**'))
    )
    r3 = _row([
        _cell(disb_cell_xml, COL_LABEL, valign='top'),
        _cell(_para(_run(disb_fmt)), COL_VALUE),
        info_cont,
    ], height_dxa=1900)

    r4 = _row([
        _cell(_para(_run('Total Amount to be Paid to Us', bold=True)), COL_LABEL),
        _cell(_para(_run(pa_fmt)), COL_VALUE),
        info_cont,
    ], height_dxa=1440)

    tdc_cell_xml = (
        _para(_run('Total Dollar Cost', bold=True))
        + _para(_run('Total Amount to be Paid to Us -'))
        + _para(_run('Total Amount of the Commercial Financing'))
    )
    r5 = _row([
        _cell(tdc_cell_xml, COL_LABEL),
        _cell(_para(_run(tdc_fmt)), COL_VALUE),
        info_cont,
    ], height_dxa=1440)

    top_table = _table([r1, r2, r3, r4, r5])

    # ── Payment + prepayment table ──────────────────────────────────────────
    L2, R2 = 2340, 8460

    if is_weekly:
        pay_plain = ('We will debit your business bank account once per week, on your '
                     'scheduled payment day. ')
    else:
        pay_plain = ('We will debit your business bank account each business day '
                     '(Monday-Friday). ')
    pay_italic = ('If a debit is scheduled for a bank holiday, the payment will be debited '
                  'the next business day. ')
    pay_tail = ('For details on your right to adjust the payment amount, see Section 3 of '
                'your contract.')

    pay_cell = (
        _para(_run(f'{pmt_fmt} / {per_label}'))
        + _empty_para(size=12)
        + _para([_run(pay_plain), _run(pay_italic, italic=True), _run(pay_tail)])
    )
    pr1 = _row([
        _cell(_para(_run('Initial Estimated', bold=True))
              + _para(_run('Payment Amount', bold=True)), L2),
        _cell(pay_cell, R2, valign='top'),
    ], height_dxa=1100)

    prepay1 = ('If you pay off the financing faster than required, you will/ will not be '
               'required to pay additional fees.')
    prepay2 = ('If you pay off the financing faster than required, you will not receive a '
               'discount for prepayment. Unless there was an Early Pay-Off Addendum to your '
               'contract executed at the time of the Agreement. Any other discount is at the '
               'sole discretion of the Provider')

    pr2 = _row([
        _cell(_para(_run('Description of', bold=True))
              + _para(_run('Prepayment Policies', bold=True)), L2, vmerge='restart'),
        _cell(_para(_run(prepay1)), R2, valign='top'),
    ], height_dxa=560)
    pr3 = _row([
        _cell(_empty_para(), L2, vmerge='continue'),
        _cell(_para(_run(prepay2)), R2, valign='top'),
    ], height_dxa=860)

    bottom_table = _table([pr1, pr2, pr3])

    # ── Acknowledgment + signature ──────────────────────────────────────────
    ack = _para(_run('By signing below, you acknowledge that you have received a copy of '
                     'this disclosure form.'), spacing_before=240, spacing_after=240)

    sig_line = '_' * 46
    date_line = '_' * 22
    sig = _para([_run('Signature of Merchant:  ', bold=True), _run(sig_line)],
                spacing_before=120, spacing_after=240)
    dte = _para([_run('Date:  ', bold=True), _run(date_line)], spacing_after=120)

    sect_pr = ('<w:sectPr>'
               '<w:pgSz w:w="12240" w:h="15840"/>'
               '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
               'w:header="720" w:footer="720" w:gutter="0"/>'
               '</w:sectPr>')

    body = title + top_table + bottom_table + ack + sig + dte + sect_pr

    document_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<w:document {NS}>\n'
                    f'<w:body>{body}</w:body>\n'
                    '</w:document>')

    # ── DOCX scaffolding ────────────────────────────────────────────────────
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
