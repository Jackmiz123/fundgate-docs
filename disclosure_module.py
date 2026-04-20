"""
Disclosure module — builds state CFDL disclosures as DOCX bytes.
XML generated directly to match the GA sample disclosure exactly.
"""
import io, zipfile, re
from docx import Document
from docx.shared import Pt, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime

# ── State config ───────────────────────────────────────────────────────────────
DISCLOSURE_STATES = {
    'FL': {
        'name': 'Florida',
        'statute': 'Florida Statutes §§559.961\u2013559.9615 (Florida Commercial Financing Disclosure Law, eff. January 1, 2024)',
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
        'statute': 'Louisiana R.S. 9:3578.1 et seq. (Louisiana Commercial Financing Disclosure Law)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': False,
    },
    'KS': {
        'name': 'Kansas',
        'statute': 'Kansas SB 345 \u2014 Commercial Financing Disclosure Act (eff. July 1, 2024)',
        'not_loan': 'This transaction is a purchase and sale of future receivables and is NOT a loan. Amounts charged are NOT interest.',
        'kansas_labels': True,
    },
    'MO': {
        'name': 'Missouri',
        'statute': 'Missouri Revised Statutes \u00a7427.300 et seq. (Commercial Financing Disclosure Law, eff. February 28, 2025)',
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

# ── XML helpers ────────────────────────────────────────────────────────────────
FONT = 'Arial'
NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'

def _rpr(bold=False, italic=False, sz=20):
    b = '<w:b/><w:bCs/>' if bold else '<w:b w:val="false"/><w:bCs w:val="false"/>'
    i = '<w:i/><w:iCs/>' if italic else '<w:i w:val="false"/><w:iCs w:val="false"/>'
    return (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
            f'{b}{i}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')

def _ppr(before=0, after=40, jc='left', indent=None, indent_lr=None):
    ind = ''
    if indent_lr:
        ind = f'<w:ind w:left="{indent_lr}" w:right="{indent_lr}"/>'
    elif indent:
        ind = f'<w:ind w:left="{indent}"/>'
    return f'<w:pPr><w:spacing w:before="{before}" w:after="{after}"/><w:jc w:val="{jc}"/>{ind}</w:pPr>'

def _run(text, bold=False, italic=False, sz=20):
    # Escape XML special chars
    text = text.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')
    return f'<w:r>{_rpr(bold,italic,sz)}<w:t xml:space="preserve">{text}</w:t></w:r>'

def _para(runs, before=0, after=40, jc='left', indent=None, indent_lr=None):
    return f'<w:p>{_ppr(before,after,jc,indent,indent_lr)}{"".join(runs)}</w:p>'

TC_BORDERS = ('<w:tcBorders>'
              '<w:top w:val="single" w:color="000000" w:sz="4"/>'
              '<w:left w:val="single" w:color="000000" w:sz="4"/>'
              '<w:bottom w:val="single" w:color="000000" w:sz="4"/>'
              '<w:right w:val="single" w:color="000000" w:sz="4"/>'
              '</w:tcBorders>')

def _tcpr(w, span=1):
    return (f'<w:tcPr><w:tcW w:type="dxa" w:w="{w}"/><w:gridSpan w:val="{span}"/>'
            f'{TC_BORDERS}'
            f'<w:shd w:fill="FFFFFF" w:val="clear"/>'
            f'<w:tcMar><w:top w:type="dxa" w:w="120"/><w:left w:type="dxa" w:w="100"/>'
            f'<w:bottom w:type="dxa" w:w="120"/><w:right w:type="dxa" w:w="100"/></w:tcMar>'
            f'<w:vAlign w:val="top"/></w:tcPr>')

TBL_BORDERS = ('<w:tblBorders>'
               '<w:top w:val="single" w:color="auto" w:sz="4"/>'
               '<w:left w:val="single" w:color="auto" w:sz="4"/>'
               '<w:bottom w:val="single" w:color="auto" w:sz="4"/>'
               '<w:right w:val="single" w:color="auto" w:sz="4"/>'
               '<w:insideH w:val="single" w:color="auto" w:sz="4"/>'
               '<w:insideV w:val="single" w:color="auto" w:sz="4"/>'
               '</w:tblBorders>')

def _tbl(grid_cols, rows_xml, w=11520):
    grid = ''.join(f'<w:gridCol w:w="{c}"/>' for c in grid_cols)
    rows = ''.join(rows_xml)
    return (f'<w:tbl>'
            f'<w:tblPr><w:tblW w:type="dxa" w:w="{w}"/>'
            f'{TBL_BORDERS}</w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid>'
            f'{rows}</w:tbl>')

def _tc(w, paras_xml, span=1):
    return f'<w:tc>{_tcpr(w,span)}{"".join(paras_xml)}</w:tc>'

def _tr(*cells):
    return f'<w:tr>{"".join(cells)}</w:tr>'

def _sig_line_para():
    """A paragraph with bottom border as the signature line."""
    return ('<w:p><w:pPr>'
            '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'
            f'<w:spacing w:before="200" w:after="40"/>'
            '</w:pPr>'
            f'<w:r>{_rpr(sz=18)}<w:t xml:space="preserve"> </w:t></w:r></w:p>')

# ── Main function ──────────────────────────────────────────────────────────────
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
    weekly_amt = _fmt_currency(_n(data, 'Specific_Weekly_Amount'))
    ach_freq   = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()

    signer1_name  = (data.get('Owner_Guarantor_1', '') or '').title()
    signer1_title = (data.get('Title', '') or '').title()
    signer2_name  = (data.get('Owner_Guarantor_2', '') or '').title() if two_signers else ''
    signer2_title = (data.get('Title_2', '') or '').title() if two_signers else ''

    kansas = cfg.get('kansas_labels', False)
    freq_word = 'Business Week' if 'week' in ach_freq else 'Business Day'
    freq_checkbox = (
        '\u2612Every Business Week  (i.e., one debit per week on a designated business day, '
        'excluding bank holidays. Payments scheduled for a bank holiday will be debited the next '
        'business day with the regular payment)'
        if 'week' in ach_freq else
        '\u2612Every Business Day (i.e., Monday through Friday, excluding bank holidays. Payments '
        'scheduled for a bank holiday will be debited the next business day with the regular payment)'
    )

    initial_payment = _fmt_currency(_n(data, 'Specific_Weekly_Amount') if 'week' in ach_freq
                                    else _n(data, 'Specific_Daily_Amount'))

    # ── Build XML body ─────────────────────────────────────────────────────────

    # Title
    title_xml = _para([_run(f"{cfg['name'].upper()} COMMERCIAL FINANCING DISCLOSURE",
                            bold=True, sz=22)],
                      before=0, after=100, jc='center')

    # Date
    date_xml = _para([_run('Disclosure Date: ', sz=20),
                      _run(date_display, bold=True, sz=20)],
                     before=0, after=80, jc='right')

    # ── Table 0: Header ─────────────────────────────────────────────────────
    left_cell_paras = [
        _para([_run(f'Recipient: {merchant_name}', bold=True)], after=40),
        _para([_run(f'DBA: {merchant_dba}', bold=True)], after=40),
        _para([_run(f'Address: {address}', bold=True)], after=40),
    ]
    right_cell_paras = [
        _para([_run('Provider', bold=True)], after=40),
        _para([_run('Name: FundGate LLC', bold=True)], after=40),
        _para([_run('Address: 1202 Avenue U, Suite 1175, Brooklyn NY 11229', bold=True)], after=40),
        _para([_run('Phone Number: 929-256-7464', bold=True)], after=40),
        _para([_run('E-mail Address: admin@fundgatellc.com', bold=True)], after=40),
    ]
    desc_para = _para(
        [_run(f'This Commercial Financing Disclosure is being provided to the Recipient ("you") by the '
              f'Provider ("we"or"us") as required by law and is dated as of the Disclosure Date.',
              italic=True)],
        before=0, after=0
    )

    tbl0 = _tbl([5760, 5760], [
        _tr(_tc(5760, left_cell_paras), _tc(5760, right_cell_paras)),
        _tr(_tc(11520, [desc_para], span=2)),
    ])

    # ── Table 1: Amounts ────────────────────────────────────────────────────
    if kansas:
        amounts_rows = [
            _tr(_tc(9048, [_para([_run('1.  Total Amount of Funds Provided', bold=True)], after=40)]),
                _tc(2472, [_para([_run(pp_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [
                _para([_run('2.  Total Amount of Funds Disbursed', bold=True)], after=40),
                _para([_run(f'   Fees deducted or withheld at disbursement \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  {orig_fmt}', sz=20)], after=40),
                _para([_run(f'   Amount deducted for prior balance paid to us \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  $0.00', sz=20)], after=40),
                _para([_run(f'   Amount deducted and paid to third parties on your behalf \u2026\u2026  $0.00', sz=20)], after=40),
            ]),
                _tc(2472, [_para([_run(dis_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [_para([_run('3.  Total of Payments', bold=True)], after=40)]),
                _tc(2472, [_para([_run(pa_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [_para([_run('4.  Total Dollar Cost of Financing', bold=True)], after=40)]),
                _tc(2472, [_para([_run(cost_fmt, bold=True)], after=20, jc='right')])),
        ]
    else:
        amounts_rows = [
            _tr(_tc(9048, [_para([_run('1.  Total Amount of Funding Provided', bold=True)], after=40)]),
                _tc(2472, [_para([_run(pp_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [
                _para([_run('2.  Amounts Deducted from Funding Provided', bold=True)], after=40),
                _para([_run(f'   Fees deducted or withheld at disbursement \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  {orig_fmt}', sz=20)], after=40),
                _para([_run(f'   Amount deducted for prior balance paid to us \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  $0.00', sz=20)], after=40),
                _para([_run(f'   Amount deducted and paid to third parties on your behalf \u2026\u2026  $0.00', sz=20)], after=40),
            ]),
                _tc(2472, [_para([_run(orig_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [_para([_run('3.  Total Amount of Funds Disbursed (1 minus 2)', bold=True)], after=40)]),
                _tc(2472, [_para([_run(dis_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [_para([_run('4.  Total Amount to be Paid to Us', bold=True)], after=40)]),
                _tc(2472, [_para([_run(pa_fmt, bold=True)], after=20, jc='right')])),
            _tr(_tc(9048, [_para([_run('5.  Total Dollar Cost (4 minus 1)', bold=True)], after=40)]),
                _tc(2472, [_para([_run(cost_fmt, bold=True)], after=20, jc='right')])),
        ]

    tbl1 = _tbl([9048, 2472], amounts_rows)

    # ── Table 2: Payment / prepayment ───────────────────────────────────────
    payment_paras = [
        _para([_run('We will collect the Total Amount to be Paid to Us by debiting your business bank '
                    'account in periodic installments or "payments" that will occur with the following frequency:')],
              after=40),
        _para([_run(freq_checkbox)], after=40),
        _para([_run(f'The initial payment will be '),
               _run(f'{initial_payment}.', bold=True),
               _run(f' We based your initial payment on '),
               _run(spec_pct if '%' in spec_pct else f'{spec_pct}%', bold=True),
               _run(f' of your estimated sales revenue. For details on your right to adjust any payment amount, '
                    f'see Section 3 of your Purchase Agreement.')],
              after=0),
    ]

    prepay_text = (f'If you pay off the financing faster than required, you may pay a reduced amount per the '
                   f'Addendum to Merchant Cash Advance Agreement dated {date_display}, which sets forth the '
                   f'contractual rights of the parties related to prepayment. No additional fees will be charged for prepayment.')

    tbl2 = _tbl([2869, 8651], [
        _tr(_tc(2869, [_para([_run('Manner, frequency, and amount of each payment', bold=True)], after=0)]),
            _tc(8651, payment_paras)),
        _tr(_tc(2869, [_para([_run('Description of Prepayment Policies', bold=True)], after=0)]),
            _tc(8651, [_para([_run(prepay_text)], after=0)])),
    ])

    # ── Acknowledgment ──────────────────────────────────────────────────────
    ack_xml = _para([_run('By signing below, you acknowledge that you have received a copy of this disclosure form.')],
                    before=80, after=80)

    # ── Signature table ─────────────────────────────────────────────────────
    def _sig_line_xml():
        return ('<w:p><w:pPr>'
                '<w:pBdr><w:bottom w:val="single" w:sz="6" w:color="000000" w:space="1"/></w:pBdr>'
                '<w:spacing w:before="200" w:after="40"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'
                '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
                '<w:b w:val="0"/><w:i w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
                '<w:t xml:space="preserve"> </w:t></w:r></w:p>')

    def _label_xml(text):
        return ('<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="60"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'
                '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
                '<w:b w:val="0"/><w:i w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
                f'<w:t>{text.replace("&","&amp;").replace("<","&lt;")}</w:t></w:r></w:p>')

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

    # Build sig column content
    s1_label = f'Recipient Signature — {signer1_name}, {signer1_title}' if signer1_title else f'Recipient Signature — {signer1_name}'
    sig_col_content = _sig_line_xml() + _label_xml(s1_label)

    date_col_content = _sig_line_xml() + _label_xml('Date')

    if two_signers and signer2_name:
        s2_label = f'Recipient Signature — {signer2_name}, {signer2_title}' if signer2_title else f'Recipient Signature — {signer2_name}'
        sig_col_content += _spacer_xml() + _sig_line_xml() + _label_xml(s2_label)
        date_col_content += _spacer_xml() + _sig_line_xml() + _label_xml('Date')

    tbl3 = (
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
        '<w:tblGrid><w:gridCol w:w="5628"/><w:gridCol w:w="662"/><w:gridCol w:w="5230"/></w:tblGrid>'
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="5628" w:type="dxa"/>{NO_BORDER_TC}</w:tcPr>{sig_col_content}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="662" w:type="dxa"/>{NO_BORDER_TC}</w:tcPr>'
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p></w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="5230" w:type="dxa"/>{NO_BORDER_TC}</w:tcPr>{date_col_content}</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )

    # ── Footer ──────────────────────────────────────────────────────────────
    footer_xml = _para(
        [_run(f"Pursuant to {cfg['statute']}. {cfg['not_loan']}", italic=True, sz=18)],
        before=80, after=0, jc='center'
    )

    # ── Assemble document XML ───────────────────────────────────────────────
    body_content = (title_xml + date_xml + tbl0 + tbl1 + tbl2 +
                    ack_xml + tbl3 + footer_xml)

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {NS}>'
        '<w:body>'
        + body_content +
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
        '<w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )

    # ── Package as DOCX ─────────────────────────────────────────────────────
    # Use the GA sample as the base DOCX (for styles, fonts, etc)
    import os
    sample_path = os.path.join(os.path.dirname(__file__), 'disclosure_sample.docx')
    if not os.path.exists(sample_path):
        # Fall back to creating from scratch
        base_path = None
    else:
        base_path = sample_path

    # Create minimal DOCX from scratch
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        # Minimal required DOCX files
        zout.writestr('word/document.xml', doc_xml.encode('utf-8'))
        zout.writestr('[Content_Types].xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>')
        zout.writestr('_rels/.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
        zout.writestr('word/_rels/document.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>')
    return buf.getvalue()
