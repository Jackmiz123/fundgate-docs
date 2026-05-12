#!/usr/bin/env python3
"""FundGate Weekly Contract Generator — Render.com deployment"""
import http.server, json, zipfile, re, os, subprocess, tempfile, shutil, io
from http.server import HTTPServer, BaseHTTPRequestHandler
from disclosure_module import build_disclosure_bytes
from payoff_module import build_payoff_letter
from zero_balance_module import build_zero_balance_letter
from ca_disclosure_module import build_ca_disclosure_bytes

TEMPLATE_WEEKLY = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_WEEKLY.docx')
TEMPLATE_DAILY  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_DAILY.docx')
# CA-only templates (Fundkey LLC branding) — used ONLY by /generate/ca routes
TEMPLATE_WEEKLY_CA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDKEY_TEMPLATE_WEEKLY.docx')
TEMPLATE_DAILY_CA  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDKEY_TEMPLATE_DAILY.docx')
FORM            = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fundgate_form.html')

SIGNER2_BLOCKS_DIR = os.path.dirname(os.path.abspath(__file__))

FIELDS = {
    '«Agreement_Date»':             'Agreement_Date',
    '«Agreement_Date1»':            'Agreement_Date',
    '«Merchant_Legal_Name»':        'Merchant_Legal_Name',
    '«Merchant_DBA»':               'Merchant_DBA',
    '«Entity_Type»':                'Entity_Type',
    '«State_of_Organization»':      'State_of_Organization',
    '«Executive_Office_Address»':   'Executive_Office_Address',
    '«Mailing_Address»':            'Mailing_Address',
    '«Business_Start_Date»':        'Business_Start_Date',
    '«Federal_EIN»':                'Federal_EIN',
    '«Business_Phone»':             'Business_Phone',
    '«Purchase_Price»':             'Purchase_Price',
    '«Purchased_Amount»':           'Purchased_Amount',
    '«Specified_Percentage»':       'Specified_Percentage',
    '«ACH_Frequency»':              'ACH_Frequency',
    '«Specific_Weekly_Amount»':     'Specific_Weekly_Amount',
    '«Specific_Daily_Amount»':      'Specific_Daily_Amount',
    '«ACH_Program_Fee_Percentage»': 'ACH_Program_Fee_Percentage',
    '«Originiation_Fee_Percentage»':'Origination_Fee_Percentage',
    '«Merchant__1»':                'Merchant_1',
    '«Owner_Guarantor_1»':          'Owner_Guarantor_1',
    '«Guarantor_SSN»':              'Guarantor_SSN',
    '«Guarantor_Driver_License»':   'Guarantor_Driver_License',
    '«Bank_Name»':                  'Bank_Name',
    '«Routing_Number»':             'Routing_Number',
    '«Account_Number»':             'Account_Number',
    '«Authorized_Signer_Name»':     'Authorized_Signer_Name',
    '«Repurchase_30_Day_Amount»':   'Repurchase_30_Day_Amount',
    # Repurchase_31_60_Day_Amount handled dynamically via REPURCHASE_MID_BLOCK
    '«After_60_Day_Amount»':        'After_60_Day_Amount',
    '«Title»':                      'Title',
}

# Signer 2 fields (only used when twoSigners=true)
FIELDS_S2 = {
    '«Owner_Guarantor_2»':   'Owner_Guarantor_2',
    '«Title_2»':             'Title_2',
    '«Guarantor_SSN_2»':     'Guarantor_SSN_2',
    '«Guarantor_DL_2»':      'Guarantor_DL_2',
    '«Merchant_Legal_Name»': 'Merchant_Legal_Name',
    '«Merchant_DBA»':        'Merchant_DBA',
}

def load_signer2_block(filename):
    path = os.path.join(SIGNER2_BLOCKS_DIR, filename)
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    return ''

def merge_disclosure_into_contract(contract_bytes, disclosure_bytes):
    """Prepend disclosure DOCX pages to contract DOCX."""
    with zipfile.ZipFile(io.BytesIO(disclosure_bytes)) as dz:
        disc_xml = dz.read('word/document.xml').decode('utf-8')

    with zipfile.ZipFile(io.BytesIO(contract_bytes)) as cz:
        contract_files = {n: cz.read(n) for n in cz.namelist()}
        contract_xml   = contract_files['word/document.xml'].decode('utf-8')

    disc_body_match = re.search(r'<w:body>(.*)</w:body>', disc_xml, re.DOTALL)
    if not disc_body_match:
        return contract_bytes

    disc_body = disc_body_match.group(1)
    disc_body = re.sub(r'<w:sectPr\b.*?</w:sectPr>\s*$', '', disc_body, flags=re.DOTALL).rstrip()
    disc_body = re.sub(r'(<w:p[^>]*>\s*<w:pPr[^/]*/>\s*</w:p>\s*)+$', '', disc_body).rstrip()
    disc_body = re.sub(r'(<w:p[^>]*>\s*</w:p>\s*)+$', '', disc_body).rstrip()

    page_break_para = (
        '<w:p><w:r><w:lastRenderedPageBreak/><w:br w:type="page"/></w:r></w:p>\n'
    )
    disc_body = disc_body + '\n' + page_break_para

    contract_xml = contract_xml.replace('<w:body>', '<w:body>' + disc_body, 1)
    contract_xml = contract_xml.replace('<w:pgNumType w:start="1"/>', '<w:pgNumType w:start="2"/>', 1)

    contract_files['word/document.xml'] = contract_xml.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data_bytes in contract_files.items():
            zout.writestr(name, data_bytes)
    return buf.getvalue()


def fill_docx(data):
    is_daily = data.get('dealType') == 'daily'
    # CA-only flag — when True, use Fundkey-branded templates instead of FundGate.
    # All existing (non-CA) routes leave isCA absent, so this is a pure addition
    # and does not affect the FundGate Weekly/Daily/Payoff/ZeroBalance flows.
    is_ca = bool(data.get('isCA', False))
    if is_ca:
        template = TEMPLATE_DAILY_CA if is_daily else TEMPLATE_WEEKLY_CA
    else:
        template = TEMPLATE_DAILY if is_daily else TEMPLATE_WEEKLY
    with zipfile.ZipFile(template, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    doc = files['word/document.xml'].decode('utf-8')
    two_signers = data.get('twoSigners', False)

    def safe(val):
        return (val or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

    # Determine the rId that points to the Sign Here arrow image (media/image2.png)
    # in the current template — s2_block files reference rId10 (weekly's mapping) by
    # default, so daily needs a rewrite since daily maps image2.png to rId9.
    rels_xml = files['word/_rels/document.xml.rels'].decode('utf-8')
    _arrow_rid = None
    for m in re.finditer(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml):
        if m.group(2) == 'media/image2.png':
            _arrow_rid = m.group(1)
            break

    # ── Handle signer 2 placeholder blocks ──────────────────────────────────────
    if two_signers:
        def fill_s2_block(block_xml):
            for placeholder, key in FIELDS_S2.items():
                val = data.get(key, '') or ''
                if key in ('Owner_Guarantor_2', 'Merchant_Legal_Name', 'Merchant_DBA'):
                    val = val.upper()
                block_xml = block_xml.replace(placeholder, safe(val))
            # Patch rId10 references (Sign Here arrow) for templates where
            # image2.png maps to a different rId (e.g. daily uses rId9).
            if _arrow_rid and _arrow_rid != 'rId10':
                block_xml = block_xml.replace('r:embed="rId10"', f'r:embed="{_arrow_rid}"')
            return block_xml

        block_p4   = fill_s2_block(load_signer2_block('s2_block_p4.xml'))
        block_ach  = fill_s2_block(load_signer2_block('s2_block_ach.xml'))
        block_bank = fill_s2_block(load_signer2_block('s2_block_bank.xml'))
        block_add  = fill_s2_block(load_signer2_block('s2_block_add.xml'))
        block_p15  = fill_s2_block(load_signer2_block('s2_block_p15.xml'))
    else:
        block_p4 = block_ach = block_bank = block_add = block_p15 = ''

    # Replace the placeholder tokens with actual block XML (or empty string)
    doc = doc.replace('«SIGNER2_BLOCK_P4»',        block_p4)
    if not two_signers:
        block_ach = ('<w:p w:rsidR="00D3482D" w:rsidRPr="00EC45D9" '
                     'w:rsidRDefault="00D3482D" w:rsidP="00D3482D">'
                     '<w:pPr><w:pStyle w:val="BodyText"/>'
                     '<w:spacing w:before="222"/><w:ind w:left="120"/></w:pPr>'
                     '<w:r><w:rPr><w:color w:val="010202"/>'
                     '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                     '<w:t>[Attached Voided Check Here]</w:t></w:r></w:p>')
    doc = doc.replace('«SIGNER2_BLOCK_ACH»',        block_ach)
    doc = doc.replace('«SIGNER2_BLOCK_BANKLOGIN»',  block_bank)
    doc = doc.replace('«SIGNER2_BLOCK_ADDENDUM»',   block_add)

    # ── Consumer Notice page (between Bank Login and Addendum) ───────────────
    consumer_notice = (
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="120"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        '<w:t>CONSUMER NOTICE REGARDING DEBT-RELIEF COMPANIES</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="120"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:i/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        '<w:t>For informational purposes only; not a contract and not legally binding.</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="120" w:after="120"/><w:jc w:val="both"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        '<w:t xml:space="preserve">Some of the information in this guide is based, in part, on publicly available consumer-protection guidance from the Florida State Attorney General\u2019s Office website and other consumer-protection sources regarding debt-relief and debt-settlement companies. This notice is intended to help protect customers from misleading or unauthorized third-party companies that may contact you after entering into a financial agreement.</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="200" w:after="80"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
        '<w:t>Be Aware of Third-Party Debt-Relief Companies!</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="80"/><w:jc w:val="both"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        '<w:t>Certain companies advertise \u201cdebt relief,\u201d \u201cdebt settlement,\u201d or \u201cnegotiation services.\u201d Consumer-protection agencies have warned that many such companies:</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(a) Make exaggerated or unrealistic promises</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(b) Charge large upfront or recurring fees</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(c) Provide little or no real benefit</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(d) Use aggressive, misleading, or pressure-based tactics</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="80"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(e) Encourage actions that may lead to financial harm</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="80" w:after="120"/><w:jc w:val="both"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>These companies may contact you unexpectedly and may suggest that they can act on your behalf without your involvement.</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="200" w:after="80"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>Common Red Flags</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="80"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Exercise caution if a company:</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(a) Promises to \u201celiminate\u201d or drastically reduce your obligations</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(b) Guarantees results that sound too good to be true</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(c) Requests upfront or monthly fees for \u201crepresentation\u201d</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(d) Uses pressure-based sales tactics</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="80"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(e) Reaches out unexpectedly, claiming they can intervene for you</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="200" w:after="80"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>If You Have Questions or Need Assistance</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="80"/><w:jc w:val="both"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>If you ever have questions about your account, experience payment issues, or need clarification on anything, you can contact us directly. There is no need to involve any outside company:</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="40" w:after="40"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Phone: 929-256-7464</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="80"/><w:ind w:left="720"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Email: admin@fundgatellc.com</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="200" w:after="80"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>Purpose of This Notice</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:spacing w:before="0" w:after="120"/><w:jc w:val="both"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>This document is not part of your contract, is not legal advice, and is not legally binding. It is purely informational and is intended to help customers avoid misleading or unauthorized third-party companies.</w:t></w:r></w:p>'
    )
    # Replace the CONSUMER_NOTICE_PAGE token with page break + notice.
    # The token sits inside: <w:r><w:t>«CONSUMER_NOTICE_PAGE»</w:t></w:r></w:p>
    # We replace that entire tail of the paragraph, close it, add page break,
    # then add all notice paragraphs.
    doc = doc.replace(
        '<w:t>\u00abCONSUMER_NOTICE_PAGE\u00bb</w:t></w:r></w:p>',
        '<w:t></w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr>'
        '<w:r><w:br w:type="page"/></w:r></w:p>\n'
        + consumer_notice
    )

    # ── CA-only: swap the FundGate email line in the consumer notice ─────────
    # Fundkey email = admin@fundkey.com
    if is_ca:
        doc = doc.replace(
            '<w:t>Email: admin@fundgatellc.com</w:t>',
            '<w:t>Email: admin@fundkey.com</w:t>'
        )

    # ── Repurchase mid-block (3-tier vs 4-tier) ────────────────────────────
    # Build paragraphs that mirror the template's "If within thirty (30)..." styling:
    # each word as its own run, bold on the day numbers (thirty-one, (31), and,
    # forty-five, (45), days), bold + underline on the dollar amount.
    use_4tier = bool(data.get('use_4tier_repurchase'))
    r31_45 = data.get('Repurchase_31_45_Day_Amount', '') or ''
    r46_60 = data.get('Repurchase_46_60_Day_Amount', '') or ''
    r31_60 = data.get('Repurchase_31_60_Day_Amount', '') or ''

    def _run_plain(text):
        return ('<w:r><w:rPr><w:color w:val="010202"/></w:rPr>'
                '<w:t xml:space="preserve">' + safe(text) + '</w:t></w:r>')

    def _run_bold(text):
        return ('<w:r><w:rPr><w:b/><w:color w:val="010202"/></w:rPr>'
                '<w:t xml:space="preserve">' + safe(text) + '</w:t></w:r>')

    def _run_bold_underline(text):
        return ('<w:r><w:rPr><w:b/><w:color w:val="010202"/>'
                '<w:spacing w:val="-2"/>'
                '<w:u w:val="single" w:color="010202"/></w:rPr>'
                '<w:t xml:space="preserve">' + safe(text) + '</w:t></w:r>')

    def _build_mid_para(low_word, low_num, high_word, high_num, amount):
        """Build a properly styled 'If between X and Y days...' paragraph."""
        # Paragraph properties match the template's addendum body paragraphs
        ppr = ('<w:pPr><w:spacing w:before="222"/>'
               '<w:ind w:left="120"/>'
               '<w:rPr><w:b/></w:rPr></w:pPr>')
        runs = []
        runs.append(_run_plain('If between '))
        runs.append(_run_bold(low_word))
        runs.append(_run_bold(' ('))
        runs.append(_run_bold(low_num))
        runs.append(_run_bold(') '))
        runs.append(_run_bold('and'))
        runs.append(_run_bold(' '))
        runs.append(_run_bold(high_word))
        runs.append(_run_bold(' ('))
        runs.append(_run_bold(high_num))
        runs.append(_run_bold(') '))
        runs.append(_run_bold('days'))
        runs.append(_run_plain(' of the Payment of the Purchase price by the Purchaser, '
                               'the Repurchase Amount shall be '))
        runs.append(_run_bold_underline(amount))
        runs.append(_run_plain(' less any payments made prior to the payment of the '
                               'Repurchase Amount.'))
        return ('<w:p w:rsidR="00D3482D" w:rsidRPr="00EC45D9" '
                'w:rsidRDefault="00D3482D" w:rsidP="00D3482D">' +
                ppr + ''.join(runs) + '</w:p>')

    if use_4tier and r31_45 and r46_60:
        mid_xml = (
            _build_mid_para('thirty-one', '31', 'forty-five', '45', r31_45) +
            _build_mid_para('forty-six',  '46', 'sixty',      '60', r46_60)
        )
    else:
        mid_xml = _build_mid_para('thirty-one', '31', 'sixty', '60', r31_60)

    # Replace the ENTIRE containing paragraph, not just the token text. The
    # template has «REPURCHASE_MID_BLOCK» as the only content of a <w:t> inside
    # a <w:r> inside a <w:p>; replacing just the text would inject <w:p>
    # elements inside a <w:t> and LibreOffice silently drops the content.
    doc = re.sub(
        r'<w:p\b[^>]*>(?:(?!<w:p\b).)*?«REPURCHASE_MID_BLOCK».*?</w:p>',
        lambda m: mid_xml,
        doc,
        count=1,
        flags=re.DOTALL,
    )
    # P15 token is in its own paragraph — replace the whole paragraph to avoid blank page
    P15_PARA = ('<w:p w:rsidR="00D3482D" w:rsidRDefault="00D3482D">'
                '<w:pPr><w:pStyle w:val="TableParagraph"/></w:pPr>'
                '<w:r><w:t>«SIGNER2_BLOCK_P15»</w:t></w:r></w:p>')
    if P15_PARA in doc:
        doc = doc.replace(P15_PARA, block_p15)
    else:
        doc = doc.replace('«SIGNER2_BLOCK_P15»', block_p15)

    # ── ACH + Bank Login spacing fix for 2-signer mode ───────────────────────────
    if two_signers:
        sect_markers = [m.start() for m in re.finditer(r'<w:sectPr[ >]', doc)]
        if len(sect_markers) >= 27:
            s22_end = doc.find('</w:p>', sect_markers[22]) + 6
            s26_end = doc.find('</w:p>', sect_markers[26]) + 6
            chunk = doc[s22_end:s26_end]

            chunk = chunk.replace('w:before="228" w:line="247" w:lineRule="auto"',
                                  'w:before="80" w:line="232" w:lineRule="auto"')
            chunk = chunk.replace('w:before="227" w:line="247" w:lineRule="auto"',
                                  'w:before="80" w:line="232" w:lineRule="auto"')
            chunk = chunk.replace('w:before="224" w:line="247" w:lineRule="auto"',
                                  'w:before="80" w:line="232" w:lineRule="auto"')
            chunk = chunk.replace('w:before="226" w:line="247" w:lineRule="auto"',
                                  'w:before="80" w:line="232" w:lineRule="auto"')
            chunk = chunk.replace('<w:spacing w:before="221"/>',
                                  '<w:spacing w:before="80"/>')
            chunk = chunk.replace('<w:spacing w:before="222"/>',
                                  '<w:spacing w:before="80"/>')
            chunk = chunk.replace('<w:spacing w:before="223"/>',
                                  '<w:spacing w:before="80"/>')
            chunk = chunk.replace('<w:spacing w:before="229"/>',
                                  '<w:spacing w:before="80"/>')
            chunk = chunk.replace('w:before="222" w:line="456" w:lineRule="auto"',
                                  'w:before="80" w:line="360" w:lineRule="auto"')
            chunk = chunk.replace('w:before="228" w:line="456" w:lineRule="auto"',
                                  'w:before="80" w:line="360" w:lineRule="auto"')
            chunk = chunk.replace('w:before="229" w:line="456" w:lineRule="auto"',
                                  'w:before="80" w:line="360" w:lineRule="auto"')
            chunk = chunk.replace('w:before="230" w:line="456" w:lineRule="auto"',
                                  'w:before="80" w:line="360" w:lineRule="auto"')
            chunk = re.sub(r'(<w:pgMar[^>]*?)w:top="1120"', r'\1w:top="700"', chunk)

            # Adjust the Sign Here arrow vertical offset for the compressed spacing.
            # In 1-signer mode, offset 175895 aligns the arrow with the signature line.
            # In 2-signer mode, the reduced before-spacing (228→80) pushes the arrow
            # below the sig line.  Reduce offset so it stays aligned.
            chunk = chunk.replace(
                '<wp:posOffset>175895</wp:posOffset></wp:positionV>',
                '<wp:posOffset>60000</wp:posOffset></wp:positionV>'
            )

            doc = doc[:s22_end] + chunk + doc[s26_end:]

        # ── Addendum signature page spacing fix ─────────────────────────────────
        sect_markers = [m.start() for m in re.finditer(r'<w:sectPr[ >]', doc)]
        if len(sect_markers) >= 28:
            s27_end = doc.find('</w:p>', sect_markers[27]) + 6
            final_sect = doc.rfind('<w:sectPr')
            s28 = doc[s27_end:final_sect]
            s28_after = doc[final_sect:]
            s28 = s28.replace('w:before="228" w:line="456" w:lineRule="auto"', 'w:before="60" w:line="360" w:lineRule="auto"')
            s28 = s28.replace('w:before="229" w:line="456" w:lineRule="auto"', 'w:before="60" w:line="360" w:lineRule="auto"')
            s28 = s28.replace('w:before="228" w:line="247" w:lineRule="auto"', 'w:before="60" w:line="232" w:lineRule="auto"')
            s28 = s28.replace('w:before="229" w:line="247" w:lineRule="auto"', 'w:before="60" w:line="232" w:lineRule="auto"')
            s28 = s28.replace('<w:spacing w:before="228"/>', '<w:spacing w:before="60"/>')
            s28 = s28.replace('<w:spacing w:before="229"/>', '<w:spacing w:before="60"/>')
            s28 = s28.replace('<w:spacing w:before="203"/>', '<w:spacing w:before="1"/>')
            doc = doc[:s27_end] + s28 + s28_after

        # ── Remove addendum spacer paragraphs for 2-signer mode ─────────────────
        SPACER = ('<w:p w:rsidR="005119B8" w:rsidRDefault="005119B8" w:rsidP="00C13BC9">'
                  '<w:pPr><w:spacing w:before="122"/><w:ind w:left="119"/>'
                  '<w:rPr><w:b/><w:color w:val="010202"/><w:spacing w:val="-2"/>'
                  '</w:rPr></w:pPr></w:p>')
        doc = doc.replace(SPACER, '')

    # ── Fill all standard fields ─────────────────────────────────────────────────
    # Fields that must be UPPERCASE in the contract
    UPPERCASE_FIELDS = {
        'Merchant_Legal_Name', 'Merchant_DBA', 'Entity_Type', 'State_of_Organization',
        'Executive_Office_Address', 'Mailing_Address', 'Bank_Name',
        'Owner_Guarantor_1', 'Merchant_1', 'Authorized_Signer_Name',
        'Owner_Guarantor_2',
    }
    for field, key in FIELDS.items():
        val = data.get(key, '') or ''
        if key in UPPERCASE_FIELDS:
            val = val.upper()
        val = safe(val)
        doc = doc.replace(field, val)

    # ── Remove addendum content but KEEP signature page ──────────────────────
    include_addendum = data.get('includeAddendum', True)
    if include_addendum in (False, 'false', 0, '0', None):
        include_addendum = False
    if not include_addendum:
        # Find the Heading1 paragraph containing "ADDENDUM" in the second half
        # of the document.  Cut from the end of the sectPr paragraph before it
        # to the end of the sectPr paragraph before the signature page.
        # This removes the addendum CONTENT only (page 20) and keeps the
        # signature page (page 21) intact.
        add_match = None
        for _m in re.finditer(r'<w:p\b[^>]*>(?:(?!</w:p>).)*?</w:p>', doc, re.DOTALL):
            if _m.start() < len(doc) // 2:
                continue
            _ptxt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', _m.group()))
            if 'ADDENDUM' in _ptxt and 'Heading1' in _m.group():
                add_match = _m
                break
        # Find the signature-page heading paragraph "By their signatures..."
        sig_match = None
        if add_match is not None:
            for _m in re.finditer(r'<w:p\b[^>]*>(?:(?!</w:p>).)*?</w:p>', doc, re.DOTALL):
                if _m.start() <= add_match.end():
                    continue
                _ptxt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', _m.group()))
                if 'By their signatures' in _ptxt:
                    sig_match = _m
                    break
        if add_match is not None and sig_match is not None:
            # Cut point A: end of paragraph that contains the sectPr right
            # before the ADDENDUM heading (preserves Consumer Notice page break).
            _sect_a = doc.rfind('<w:sectPr', 0, add_match.start())
            _sect_a_close = doc.find('</w:sectPr>', _sect_a) + len('</w:sectPr>')
            _p_close_a = doc.find('</w:p>', _sect_a_close) + len('</w:p>')
            # Cut point B: end of paragraph that contains the sectPr right
            # before the signature-page heading (preserves sig page intact).
            _sect_b = doc.rfind('<w:sectPr', 0, sig_match.start())
            _sect_b_close = doc.find('</w:sectPr>', _sect_b) + len('</w:sectPr>')
            _p_close_b = doc.find('</w:p>', _sect_b_close) + len('</w:p>')
            if 0 < _p_close_a < _p_close_b:
                doc = doc[:_p_close_a] + doc[_p_close_b:]
                # Update wording on the sig page since there is no addendum:
                # "...bound by this addendum." → "...bound by this Agreement."
                doc = doc.replace(
                    '<w:t xml:space="preserve"> addendum.</w:t>',
                    '<w:t xml:space="preserve"> Agreement.</w:t>'
                )

    files['word/document.xml'] = doc.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    return buf.getvalue()

def docx_to_pdf(docx_bytes):
    tmp = tempfile.mkdtemp()
    try:
        docx_path = os.path.join(tmp, 'contract.docx')
        pdf_path  = os.path.join(tmp, 'contract.pdf')
        with open(docx_path, 'wb') as f:
            f.write(docx_bytes)

        # LibreOffice needs a writable home dir and no display in headless mode
        lo_env = os.environ.copy()
        lo_env['HOME'] = tmp
        lo_env['TMPDIR'] = tmp
        lo_env.pop('DISPLAY', None)

        result = subprocess.run(
            ['soffice', '--headless', '--norestore', '--nofirststartwizard',
             '--convert-to', 'pdf', '--outdir', tmp, docx_path],
            capture_output=True, timeout=120, env=lo_env
        )

        if os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                return f.read()

        # Log the error for debugging
        stderr = result.stderr.decode('utf-8', errors='replace')
        stdout = result.stdout.decode('utf-8', errors='replace')
        raise Exception(f'PDF conversion failed. stdout={stdout[:300]} stderr={stderr[:300]}')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def safe_filename(data, ext):
    deal = data.get('dealType', 'weekly').capitalize()
    dba  = re.sub(r'\s+', '_', data.get('Merchant_DBA') or data.get('Merchant_Legal_Name','Contract'))
    date = (data.get('Agreement_Date','') or '').replace('/','_')
    return f"FundGate_{deal}_{dba}_{date}.{ext}"

def safe_filename_ca(data, ext):
    """CA-only filename prefix — Fundkey LLC entity."""
    deal = data.get('dealType', 'weekly').capitalize()
    dba  = re.sub(r'\s+', '_', data.get('Merchant_DBA') or data.get('Merchant_Legal_Name','Contract'))
    date = (data.get('Agreement_Date','') or '').replace('/','_')
    return f"Fundkey_CA{deal}_{dba}_{date}.{ext}"

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args): pass

    def do_GET(self):
        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-Type','text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(open(FORM,'rb').read())
        else:
            self.send_response(404); self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Headers','Content-Type')
        self.end_headers()

    def do_POST(self):
        if self.path in ('/generate', '/generate/pdf'):
            want_pdf = self.path.endswith('/pdf')
            length = int(self.headers.get('Content-Length', 0))
            data = json.loads(self.rfile.read(length))
            try:
                docx_bytes = fill_docx(data)
                disc_bytes = build_disclosure_bytes(data)
                if disc_bytes:
                    docx_bytes = merge_disclosure_into_contract(docx_bytes, disc_bytes)
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime      = 'application/pdf'
                    fname     = safe_filename(data, 'pdf')
                else:
                    out_bytes = docx_bytes
                    mime      = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname     = safe_filename(data, 'docx')
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"; filename*=UTF-8''{fname}')
                self.send_header('Content-Length', str(len(out_bytes)))
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(out_bytes)
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type','application/json')
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        elif self.path in ('/generate/ca', '/generate/ca/pdf'):
            # CA-only route — uses Fundkey-branded templates + 3-page CA disclosure.
            # Existing FundGate routes are untouched.
            want_pdf = self.path.endswith('/pdf')
            length = int(self.headers.get('Content-Length', 0))
            data = json.loads(self.rfile.read(length))
            try:
                # Force flags so fill_docx picks Fundkey templates and blanks email.
                data['isCA'] = True
                data['State_of_Organization'] = 'CA'
                docx_bytes = fill_docx(data)
                disc_bytes = build_ca_disclosure_bytes(data)
                if disc_bytes:
                    docx_bytes = merge_disclosure_into_contract(docx_bytes, disc_bytes)
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime      = 'application/pdf'
                    fname     = safe_filename_ca(data, 'pdf')
                else:
                    out_bytes = docx_bytes
                    mime      = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname     = safe_filename_ca(data, 'docx')
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"; filename*=UTF-8''{fname}')
                self.send_header('Content-Length', str(len(out_bytes)))
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(out_bytes)
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type','application/json')
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        elif self.path in ('/generate/payoff', '/generate/payoff/pdf'):
            want_pdf = self.path.endswith('/pdf')
            length = int(self.headers.get('Content-Length', 0))
            data = json.loads(self.rfile.read(length))
            try:
                docx_bytes = build_payoff_letter(data)
                if not docx_bytes:
                    raise Exception('Missing required payoff letter fields')
                merchant = re.sub(r'\s+', '_', data.get('payoff_merchant', 'Payoff'))
                date_str = (data.get('payoff_date', '') or '').replace('/', '_')
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime = 'application/pdf'
                    fname = f'FundGate_Payoff_{merchant}_{date_str}.pdf'
                else:
                    out_bytes = docx_bytes
                    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname = f'FundGate_Payoff_{merchant}_{date_str}.docx'
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"; filename*=UTF-8''{fname}')
                self.send_header('Content-Length', str(len(out_bytes)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(out_bytes)
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        elif self.path in ('/generate/zerobalance', '/generate/zerobalance/pdf'):
            want_pdf = self.path.endswith('/pdf')
            length = int(self.headers.get('Content-Length', 0))
            data = json.loads(self.rfile.read(length))
            try:
                docx_bytes = build_zero_balance_letter(data)
                if not docx_bytes:
                    raise Exception('Missing required zero balance letter fields')
                merchant = re.sub(r'\s+', '_', data.get('zb_merchant', 'ZeroBalance'))
                date_str = (data.get('zb_date', '') or '').replace('/', '_')
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime = 'application/pdf'
                    fname = f'FundGate_ZeroBalance_{merchant}_{date_str}.pdf'
                else:
                    out_bytes = docx_bytes
                    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname = f'FundGate_ZeroBalance_{merchant}_{date_str}.docx'
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"; filename*=UTF-8''{fname}')
                self.send_header('Content-Length', str(len(out_bytes)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(out_bytes)
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        else:
            self.send_response(404); self.end_headers()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    server = HTTPServer(('0.0.0.0', port), Handler)
    print(f'FundGate server running on port {port}')
    server.serve_forever()
