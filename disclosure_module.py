"""
Disclosure module - builds state CFDL disclosures as DOCX bytes.
XML generated directly to match Jack's edited disclosure layout.
"""
import io, zipfile, re
from docx import Document
from docx.shared import Pt, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime

# -- State config --
DISCLOSURE_STATES = {
    'FL': {
        'name': 'Florida',
        'statute': 'Florida Statutes \u00a7\u00a7559.961\u2013559.9615 (Florida Commercial Financing Disclosure Law, eff. January 1, 2024)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'GA': {
        'name': 'Georgia',
        'statute': 'Georgia SB 90 (O.C.G.A. \u00a7 10-1-393.15 et seq.)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'LA': {
        'name': 'Louisiana',
        'statute': 'Louisiana HB 470 (La. R.S. 51:3161 et seq., eff. August 1, 2025)',
        'not_loan': 'This transaction is a revenue-based financing transaction and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'KS': {
        'name': 'Kansas',
        'statute': 'Kansas SB 345, Commercial Financing Disclosure Act (eff. July 1, 2024)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': True,
    },
    'MO': {
        'name': 'Missouri',
        'statute': 'Missouri Revised Statutes \u00a7427.300 et seq. (Commercial Financing Disclosure Law, eff. February 28, 2025)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'UT': {
        'name': 'Utah',
        'statute': 'Utah Code \u00a77-27-101 et seq. (Commercial Financing Registration and Disclosure Act, eff. January 1, 2023)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'CT': {
        'name': 'Connecticut',
        'statute': 'Connecticut Public Act 23-142 (Conn. Gen. Stat. \u00a736a-870 et seq., eff. July 1, 2024)',
        'not_loan': 'This transaction is a sales-based financing and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'VA': {
        'name': 'Virginia',
        'statute': 'Virginia Code \u00a76.2-2237 et seq. (Sales-Based Financing Disclosure, eff. July 1, 2022)',
        'not_loan': 'This transaction is a sales-based financing and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'TX': {
        'name': 'Texas',
        'statute': 'Texas Finance Code Ch. 306 (HB 700, Sales-Based Financing Disclosure, eff. September 2025)',
        'not_loan': 'This transaction is a sales-based financing and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'ND': {
        'name': 'North Dakota',
        'statute': 'North Dakota Century Code Ch. 13-12 (HB 1127, Money Brokers Act, eff. August 1, 2025)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
}

def _fmt_currency(val):
    try:
        n = float(str(val).replace('$','').replace(',','').replace('%',''))
        return f'${n:,.2f}'
    except:
        return str(val)

def _fmt_date(val):
    if not val: return ''
    for fmt in ('%m/%d/%Y','%m/%d/%y','%Y-%m-%d'):
        try:
            return datetime.strptime(str(val).strip(), fmt).strftime('%B %d, %Y')
        except: pass
    return str(val)

def _n(data, key):
    try: return float(str(data.get(key,0)).replace('$','').replace(',','').replace('%',''))
    except: return 0.0

# -- XML helpers --
FONT = 'Arial'
NS = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word" '
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')

def _esc(text):
    return text.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def _rpr(bold=False, italic=False, sz=20):
    b = '<w:b/><w:bCs/>' if bold else ''
    i = '<w:i/><w:iCs/>' if italic else ''
    return (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
            f'{b}{i}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')

def _run(text, bold=False, italic=False, sz=20):
    return f'<w:r>{_rpr(bold,italic,sz)}<w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'

def _ppr(before=0, after=40, jc='left'):
    return f'<w:pPr><w:spacing w:before="{before}" w:after="{after}"/><w:jc w:val="{jc}"/></w:pPr>'

def _para(runs, before=0, after=40, jc='left'):
    return f'<w:p>{_ppr(before,after,jc)}{"".join(runs)}</w:p>'

# Cell borders and shading
TC_BORDERS = ('<w:tcBorders>'
              '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              '</w:tcBorders>')

TC_MARGIN = ('<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
             '<w:bottom w:w="120" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>')

def _tcpr(w, span=1):
    sp = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ''
    return (f'<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{sp}'
            f'{TC_BORDERS}<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>'
            f'{TC_MARGIN}<w:vAlign w:val="top"/></w:tcPr>')

def _tc(w, paras_xml, span=1):
    return f'<w:tc>{_tcpr(w,span)}{"".join(paras_xml)}</w:tc>'

def _tr(*cells):
    return f'<w:tr>{"".join(cells)}</w:tr>'

# Table-level borders
TBL_BORDERS = ('<w:tblBorders>'
               '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '</w:tblBorders>')

# Sig table helpers
def _sig_line_xml():
    return ('<w:p><w:pPr>'
            '<w:pBdr><w:bottom w:val="single" w:sz="6" w:color="000000" w:space="1"/></w:pBdr>'
            '<w:spacing w:before="200" w:after="40"/></w:pPr>'
            f'<w:r>{_rpr(sz=18)}<w:t xml:space="preserve"> </w:t></w:r></w:p>')

def _label_xml(text):
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="60"/></w:pPr>'
            f'<w:r>{_rpr(sz=18)}<w:t>{_esc(text)}</w:t></w:r></w:p>')

def _spacer_xml():
    return '<w:p><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr></w:p>'

NO_BORDER_TC = ('<w:tcBorders>'
                '<w:top w:val="none" w:sz="0" w:color="FFFFFF"/>'
                '<w:left w:val="none" w:sz="0" w:color="FFFFFF"/>'
                '<w:bottom w:val="none" w:sz="0" w:color="FFFFFF"/>'
                '<w:right w:val="none" w:sz="0" w:color="FFFFFF"/>'
                '</w:tcBorders>'
                '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
                '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>')


# -- Main function --
def build_disclosure_bytes(data):
    state_code = (data.get('State_of_Organization') or '').upper().strip()
    cfg = DISCLOSURE_STATES.get(state_code)
    if not cfg:
        return None

    two_signers   = data.get('twoSigners', False)
    merchant_name = (data.get('Merchant_Legal_Name', '') or '').upper()
    merchant_dba  = (data.get('Merchant_DBA', '') or merchant_name).upper()
    address       = (data.get('Executive_Office_Address', '') or '').upper()
    date_display  = _fmt_date(data.get('Agreement_Date', ''))

    pp       = _n(data, 'Purchase_Price')
    pa       = _n(data, 'Purchased_Amount')
    orig_pct = _n(data, 'Origination_Fee_Percentage')
    ach_pct  = _n(data, 'ACH_Program_Fee_Percentage')
    total_fee_pct = orig_pct + ach_pct
    orig_amt = round(pp * total_fee_pct / 100, 2)
    disbursed= round(pp - orig_amt, 2)
    cost     = round(pa - pp, 2)

    pp_fmt   = _fmt_currency(pp)
    pa_fmt   = _fmt_currency(pa)
    orig_fmt = _fmt_currency(orig_amt)
    dis_fmt  = _fmt_currency(disbursed)
    cost_fmt = _fmt_currency(cost)

    spec_pct   = data.get('Specified_Percentage', '')
    ach_freq   = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()

    signer1_name  = (data.get('Owner_Guarantor_1', '') or '').title()
    signer1_title = (data.get('Title', '') or '').title()
    signer2_name  = (data.get('Owner_Guarantor_2', '') or '').title() if two_signers else ''
    signer2_title = (data.get('Title_2', '') or '').title() if two_signers else ''

    kansas = cfg.get('kansas_labels', False)
    freq_checkbox = (
        '\u2612Every Business Week (i.e., one debit per week on a designated business day, '
        'excluding bank holidays. Payments scheduled for a bank holiday will be debited the next '
        'business day with the regular payment)'
        if 'week' in ach_freq else
        '\u2612Every Business Day (i.e., Monday through Friday, excluding bank holidays. Payments '
        'scheduled for a bank holiday will be debited the next business day with the regular payment)'
    )

    initial_payment = _fmt_currency(_n(data, 'Specific_Weekly_Amount') if 'week' in ach_freq
                                    else _n(data, 'Specific_Daily_Amount'))

    # -- Build XML body --

    # Title
    title_xml = _para([_run(f"{cfg['name'].upper()} COMMERCIAL FINANCING DISCLOSURE",
                            bold=True, sz=22)],
                      before=0, after=100, jc='center')

    # Date (right-aligned, "Disclosure Date: " normal + date bold)
    date_xml = _para([_run('Disclosure Date: ', sz=20),
                      _run(date_display, bold=True, sz=20)],
                     before=0, after=80, jc='right')

    # -- SINGLE MAIN TABLE (4 gridCols: 3009, 2949, 3357, 2522 = 11837) --
    G = [3009, 2949, 3357, 2522]  # grid columns
    W_TOTAL = 11837
    W_LEFT = G[0] + G[1]   # 5958 (Recipient / label col for payment rows)
    W_RIGHT = G[2] + G[3]  # 5879 (Provider)
    W_LABEL = G[0] + G[1] + G[2]  # 9315 (amounts label, spans 3)
    W_AMT = G[3]            # 2522 (amounts value)
    W_PAY_LABEL = G[0]      # 3009 (payment label, 1 col)
    W_PAY_TEXT = G[1] + G[2] + G[3]  # 8828 (payment text, spans 3)

    rows = []

    # Row 0: Header (Recipient | Provider)
    left_paras = [
        _para([_run(f'Recipient: {merchant_name}', bold=True)], after=40),
        _para([_run(f'DBA: {merchant_dba}', bold=True)], after=40),
        _para([_run(f'Address: {address}', bold=True)], after=0),
    ]
    right_paras = [
        _para([_run('Provider', bold=True)], after=40),
        _para([_run('Name: FundGate LLC', bold=True)], after=40),
        _para([_run('Address: 1202 Avenue U, Suite 1175, Brooklyn NY 11229', bold=True)], after=40),
        _para([_run('Phone: 929-355-8918', bold=True)], after=40),
        _para([_run('Email: admin@fundgatellc.com', bold=True)], after=0),
    ]
    rows.append(_tr(_tc(W_LEFT, left_paras, span=2), _tc(W_RIGHT, right_paras, span=2)))

    # Row 1: Statute description (merged all 4 cols)
    desc = (f'This Commercial Financing Disclosure is being provided to the Recipient ("you") by the '
            f'Provider ("we" or "us") as required by law and is dated as of the Disclosure Date.')
    rows.append(_tr(_tc(W_TOTAL, [_para([_run(desc, italic=True)], after=0)], span=4)))

    # Rows 2-6: Amounts (label spans 3, amount spans 1)
    def _amt_row(label, value):
        return _tr(
            _tc(W_LABEL, [_para([_run(label, bold=True)], after=40)], span=3),
            _tc(W_AMT, [_para([_run(value, bold=True)], after=20, jc='right')])
        )

    if kansas:
        rows.append(_amt_row('1.  Total Amount of Funds Provided', pp_fmt))
        rows.append(_tr(
            _tc(W_LABEL, [
                _para([_run('2.  Total Amount of Funds Disbursed', bold=True)], after=40),
                _para([_run(f'   Fees deducted or withheld at disbursement \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  {orig_fmt}')], after=40),
                _para([_run(f'   Amount deducted for prior balance paid to us \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  $0.00')], after=40),
                _para([_run(f'   Amount deducted and paid to third parties on your behalf \u2026\u2026  $0.00')], after=40),
            ], span=3),
            _tc(W_AMT, [_para([_run(dis_fmt, bold=True)], after=20, jc='right')])
        ))
        rows.append(_amt_row('3.  Total of Payments', pa_fmt))
        rows.append(_amt_row('4.  Total Dollar Cost of Financing', cost_fmt))
    else:
        rows.append(_amt_row('1.  Total Amount of Funding Provided', pp_fmt))
        rows.append(_tr(
            _tc(W_LABEL, [
                _para([_run('2.  Amounts Deducted from Funding Provided', bold=True)], after=40),
                _para([_run(f'   Fees deducted or withheld at disbursement \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  {orig_fmt}')], after=40),
                _para([_run(f'   Amount deducted for prior balance paid to us \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  $0.00')], after=40),
                _para([_run(f'   Amount deducted and paid to third parties on your behalf \u2026\u2026  $0.00')], after=40),
            ], span=3),
            _tc(W_AMT, [_para([_run(orig_fmt, bold=True)], after=20, jc='right')])
        ))
        rows.append(_amt_row('3.  Total Amount of Funds Disbursed (1 minus 2)', dis_fmt))
        rows.append(_amt_row('4.  Total Amount to be Paid to Us', pa_fmt))
        rows.append(_amt_row('5.  Total Dollar Cost (4 minus 1)', cost_fmt))

    # Payment row (1 col label, 3 col text)
    payment_paras = [
        _para([_run('We will collect the Total Amount to be Paid to Us by debiting your business bank '
                    'account in periodic installments or "payments" that will occur with the following frequency:')],
              after=40),
        _para([_run(freq_checkbox)], after=40),
        _para([_run(f'The initial payment will be '),
               _run(f'{initial_payment}.', bold=True),
               _run(f' We based your initial payment on '),
               _run(f'{spec_pct}%', bold=True),
               _run(f' of your estimated sales revenue. For details on your right to adjust any payment amount, '
                    f'see Section 3 of your Purchase Agreement.')],
              after=0),
    ]
    rows.append(_tr(
        _tc(W_PAY_LABEL, [_para([_run('Manner, frequency, and amount of each payment', bold=True)], after=0)]),
        _tc(W_PAY_TEXT, payment_paras, span=3)
    ))

    # Prepayment row
    prepay_text = (f'If you pay off the financing faster than required, you may pay a reduced amount per the '
                   f'Addendum to Merchant Cash Advance Agreement dated {date_display}, which sets forth the '
                   f'contractual rights of the parties related to prepayment. No additional fees will be charged for prepayment.')
    rows.append(_tr(
        _tc(W_PAY_LABEL, [_para([_run('Description of Prepayment Policies', bold=True)], after=0)]),
        _tc(W_PAY_TEXT, [_para([_run(prepay_text)], after=0)], span=3)
    ))

    # Assemble main table
    grid_xml = ''.join(f'<w:gridCol w:w="{c}"/>' for c in G)
    main_tbl = (f'<w:tbl>'
                f'<w:tblPr><w:tblW w:w="{W_TOTAL}" w:type="dxa"/>'
                f'<w:tblInd w:w="-80" w:type="dxa"/>'
                f'{TBL_BORDERS}'
                f'<w:tblLook w:val="04A0"/>'
                f'</w:tblPr>'
                f'<w:tblGrid>{grid_xml}</w:tblGrid>'
                f'{"".join(rows)}</w:tbl>')

    # -- Acknowledgment --
    ack_xml = _para([_run('By signing below, you acknowledge that you have received a copy of this disclosure form.')],
                    before=80, after=80)

    # -- Signature table --
    s1_label = f'Recipient Signature - {signer1_name}, {signer1_title}' if signer1_title else f'Recipient Signature - {signer1_name}'
    sig_col_content = _sig_line_xml() + _label_xml(s1_label)
    date_col_content = _sig_line_xml() + _label_xml('Date')

    if two_signers and signer2_name:
        s2_label = f'Recipient Signature - {signer2_name}, {signer2_title}' if signer2_title else f'Recipient Signature - {signer2_name}'
        sig_col_content += _spacer_xml() + _sig_line_xml() + _label_xml(s2_label)
        date_col_content += _spacer_xml() + _sig_line_xml() + _label_xml('Date')

    sig_tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="11520" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="none" w:sz="0" w:color="FFFFFF"/>'
        '<w:left w:val="none" w:sz="0" w:color="FFFFFF"/>'
        '<w:bottom w:val="none" w:sz="0" w:color="FFFFFF"/>'
        '<w:right w:val="none" w:sz="0" w:color="FFFFFF"/>'
        '<w:insideH w:val="none" w:sz="0" w:color="FFFFFF"/>'
        '<w:insideV w:val="none" w:sz="0" w:color="FFFFFF"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="5630"/><w:gridCol w:w="660"/><w:gridCol w:w="5230"/></w:tblGrid>'
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="5630" w:type="dxa"/>{NO_BORDER_TC}</w:tcPr>{sig_col_content}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="660" w:type="dxa"/>{NO_BORDER_TC}</w:tcPr>'
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p></w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="5230" w:type="dxa"/>{NO_BORDER_TC}</w:tcPr>{date_col_content}</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )

    # -- Footer --
    footer_xml = _para(
        [_run(f"Pursuant to {cfg['statute']}. {cfg['not_loan']}", italic=True, sz=18)],
        before=80, after=0, jc='center'
    )

    # -- Assemble document XML --
    body_content = title_xml + date_xml + main_tbl + ack_xml + sig_tbl + footer_xml

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {NS}>'
        '<w:body>'
        + body_content +
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
        '<w:pgMar w:top="720" w:right="180" w:bottom="720" w:left="360" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )

    # -- Package as DOCX using python-docx shell --
    shell_doc = Document()
    section = shell_doc.sections[0]
    section.left_margin = Twips(360)
    section.right_margin = Twips(180)
    shell_buf = io.BytesIO()
    shell_doc.save(shell_buf)
    shell_buf.seek(0)

    buf = io.BytesIO()
    with zipfile.ZipFile(shell_buf, 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == 'word/document.xml':
                    zout.writestr(item, doc_xml.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item))
    return buf.getvalue()
