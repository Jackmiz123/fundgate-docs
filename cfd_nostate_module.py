"""
No-state Commercial Financing Disclosure (refinance / net-amount version).

STANDALONE module. It does NOT modify or affect the state CFDL disclosures
(disclosure_module / ny / ct / ca). It only reuses the stateless XML/format
helpers from disclosure_module so the look matches the existing disclosures.

Difference vs the state version:
  * No state name in the title, no statute footer (this is the "NO STATE"
    Commercial Financing Disclosure).
  * "Amount deducted for prior balance paid to us" is a real manual entry
    (the refinance payoff) instead of a hardcoded $0.00.
  * Line 3 "Total Amount of Funds Disbursed" is recomputed as
    Funding Provided minus (fees + prior balance + third parties), so the
    merchant sees the true net amount reaching them.
"""
import io, zipfile, os
from disclosure_module import (
    _fmt_currency, _fmt_date, _n, _run, _para, _tbl, _tc, _tr, NS,
)


def build_nostate_disclosure_bytes(data):
    # ── Branding (entity-aware, mirrors the state module) ───────────────────
    is_fundkey     = bool(data.get('isFundkey', False)) or bool(data.get('isCA', False))
    provider_name  = 'Fundkey LLC' if is_fundkey else 'FundGate LLC'
    provider_email = 'admin@fundkeyllc.com' if is_fundkey else 'admin@fundgatellc.com'

    two_signers   = data.get('twoSigners', False)
    merchant_name = (data.get('Merchant_Legal_Name', '') or '').upper()
    merchant_dba  = (data.get('Merchant_DBA', '') or merchant_name).upper()
    address       = (data.get('Executive_Office_Address', '') or '').upper()
    date_display  = _fmt_date(data.get('Agreement_Date', ''))

    # ── Amounts ─────────────────────────────────────────────────────────────
    pp       = _n(data, 'Purchase_Price')        # 1. Total Amount of Funding Provided
    pa       = _n(data, 'Purchased_Amount')      # 4. Total Amount to be Paid to Us
    orig_pct = _n(data, 'Origination_Fee_Percentage')
    ach_pct  = _n(data, 'ACH_Program_Fee_Percentage')
    ach_mode  = (data.get('ACH_Program_Fee_Mode', 'pct') or 'pct').lower()
    orig_mode = (data.get('Origination_Fee_Mode', 'pct') or 'pct').lower()
    ach_amt       = ach_pct if ach_mode == 'dollar' else round(pp * ach_pct / 100, 2)
    orig_amt_only = orig_pct if orig_mode == 'dollar' else round(pp * orig_pct / 100, 2)
    fees          = round(ach_amt + orig_amt_only, 2)   # Fees deducted/withheld at disbursement

    prior_balance = _n(data, 'Prior_Balance_Amount')    # refinance payoff (manual entry)
    third_party   = _n(data, 'Third_Party_Amount')      # usually 0.00

    deducted_total = round(fees + prior_balance + third_party, 2)   # 2. total deducted
    disbursed      = round(pp - deducted_total, 2)                  # 3. net to merchant
    cost           = round(pa - pp, 2)                              # 5. total dollar cost

    pp_fmt    = _fmt_currency(pp)
    pa_fmt    = _fmt_currency(pa)
    fees_fmt  = _fmt_currency(fees)
    prior_fmt = _fmt_currency(prior_balance)
    third_fmt = _fmt_currency(third_party)
    ded_fmt   = _fmt_currency(deducted_total)
    dis_fmt   = _fmt_currency(disbursed)
    cost_fmt  = _fmt_currency(cost)

    spec_pct = data.get('Specified_Percentage', '')
    ach_freq = (data.get('ACH_Frequency', 'weekly') or 'weekly').lower()
    initial_payment = _fmt_currency(_n(data, 'Specific_Weekly_Amount') if 'week' in ach_freq
                                    else _n(data, 'Specific_Daily_Amount'))

    signer1_name  = (data.get('Owner_Guarantor_1', '') or '').title()
    signer1_title = (data.get('Title', '') or '').title()
    signer2_name  = (data.get('Owner_Guarantor_2', '') or '').title() if two_signers else ''
    signer2_title = (data.get('Title_2', '') or '').title() if two_signers else ''

    freq_checkbox = (
        '\u2612Every Business Week  (i.e., one debit per week on a designated business day, '
        'excluding bank holidays. Payments scheduled for a bank holiday will be debited the next '
        'business day with the regular payment)'
        if 'week' in ach_freq else
        '\u2612Every Business Day (i.e., Monday through Friday, excluding bank holidays. Payments '
        'scheduled for a bank holiday will be debited the next business day with the regular payment)'
    )

    # ── Title (no state) + date ─────────────────────────────────────────────
    title_xml = _para([_run("COMMERCIAL FINANCING DISCLOSURE", bold=True, sz=22)],
                      before=0, after=100, jc='center')
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
        _para([_run(f'Name: {provider_name}', bold=True)], after=40),
        _para([_run('Address: 1202 Avenue U, Suite 1175, Brooklyn NY 11229', bold=True)], after=40),
        _para([_run('Phone Number: 631-772-9020', bold=True)], after=40),
        _para([_run(f'E-mail Address: {provider_email}', bold=True)], after=40),
    ]
    desc_para = _para(
        [_run('This Commercial Financing Disclosure is being provided to the Recipient ("you") by the '
              'Provider ("we"or"us") as required by law and is dated as of the Disclosure Date.',
              italic=True)],
        before=0, after=0
    )
    tbl0 = _tbl([5760, 5760], [
        _tr(_tc(5760, left_cell_paras), _tc(5760, right_cell_paras)),
        _tr(_tc(11520, [desc_para], span=2)),
    ])

    # ── Table 1: Amounts (with real prior-balance line + net disbursed) ──────
    dots = '\u2026' * 11
    amounts_rows = [
        _tr(_tc(9048, [_para([_run('1.  Total Amount of Funding Provided', bold=True)], after=40)]),
            _tc(2472, [_para([_run(pp_fmt, bold=True)], after=20, jc='right')])),
        _tr(_tc(9048, [
            _para([_run('2.  Amounts Deducted from Funding Provided', bold=True)], after=40),
            _para([_run(f'   Fees deducted or withheld at disbursement {dots}  {fees_fmt}', sz=20)], after=40),
            _para([_run(f'   Amount deducted for prior balance paid to us \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026  {prior_fmt}', sz=20)], after=40),
            _para([_run(f'   Amount deducted and paid to third parties on your behalf \u2026\u2026  {third_fmt}', sz=20)], after=40),
        ]),
            _tc(2472, [_para([_run(ded_fmt, bold=True)], after=20, jc='right')])),
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
        _para([_run('The initial payment will be '),
               _run(f'{initial_payment}.', bold=True),
               _run(' We based your initial payment on '),
               _run(spec_pct if '%' in str(spec_pct) else f'{spec_pct}%', bold=True),
               _run(' of your estimated sales revenue. For details on your right to adjust any payment amount, '
                    'see Section 3 of your Purchase Agreement.')],
              after=0),
    ]
    prepay_text = ('If you pay off the financing faster than required, you may pay a reduced amount per the '
                   f'Addendum to Merchant Cash Advance Agreement dated {date_display}, which sets forth the '
                   'contractual rights of the parties related to prepayment. No additional fees will be charged for prepayment.')
    tbl2 = _tbl([2869, 8651], [
        _tr(_tc(2869, [_para([_run('Manner, frequency, and amount of each payment', bold=True)], after=0)]),
            _tc(8651, payment_paras)),
        _tr(_tc(2869, [_para([_run('Description of Prepayment Policies', bold=True)], after=0)]),
            _tc(8651, [_para([_run(prepay_text)], after=0)])),
    ])

    # ── Acknowledgment ──────────────────────────────────────────────────────
    ack_xml = _para([_run('By signing below, you acknowledge that you have received a copy of this disclosure form.')],
                    before=80, after=80)

    # ── Signature table (borderless, lines + labels) ────────────────────────
    def _sig_line_xml():
        return ('<w:p><w:pPr>'
                '<w:pBdr><w:bottom w:val="single" w:sz="6" w:color="000000" w:space="1"/></w:pBdr>'
                '<w:spacing w:before="200" w:after="40"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'
                '<w:b w:val="0"/><w:i w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
                '<w:t xml:space="preserve"> </w:t></w:r></w:p>')

    def _label_xml(text):
        return ('<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="60"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'
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

    s1_label = f'Recipient Signature \u2014 {signer1_name}, {signer1_title}' if signer1_title else f'Recipient Signature \u2014 {signer1_name}'
    sig_col_content = _sig_line_xml() + _label_xml(s1_label)
    date_col_content = _sig_line_xml() + _label_xml('Date')

    if two_signers and signer2_name:
        s2_label = f'Recipient Signature \u2014 {signer2_name}, {signer2_title}' if signer2_title else f'Recipient Signature \u2014 {signer2_name}'
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

    # ── Assemble (NO statute footer) ────────────────────────────────────────
    body_content = title_xml + date_xml + tbl0 + tbl1 + tbl2 + ack_xml + tbl3

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

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
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
