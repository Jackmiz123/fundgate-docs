#!/usr/bin/env python3
"""FundGate Weekly Contract Generator — Render.com deployment"""
import http.server, json, zipfile, re, os, subprocess, tempfile, shutil, io
from http.server import HTTPServer, BaseHTTPRequestHandler
from disclosure_module import build_disclosure_bytes
from cfd_nostate_module import build_nostate_disclosure_bytes
from payoff_module import build_payoff_letter
from zero_balance_module import build_zero_balance_letter
from ca_disclosure_module import build_ca_disclosure_bytes
from ny_disclosure_module import build_ny_disclosure_bytes
from ct_disclosure_module import build_ct_disclosure_bytes
from ut_disclosure_module import build_ut_disclosure_bytes
import db_module
import auth_module

# Characters that are ILLEGAL in XML 1.0 / OOXML (C0 control chars except
# tab, newline, carriage return). A stray one of these in any field value
# (commonly from copy-pasting a name or address out of a PDF/email) makes
# document.xml not well-formed, so Word reports "problems with the contents"
# even though the file downloads. Stripping them is a no-op for clean data
# and never affects legitimate contract text.
_ILLEGAL_XML_CHARS = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

def strip_illegal_xml_chars(s):
    return _ILLEGAL_XML_CHARS.sub('', s) if s else s

TEMPLATE_WEEKLY = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_WEEKLY.docx')
TEMPLATE_DAILY  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_DAILY.docx')
# CA-only templates (Fundkey LLC branding) — used ONLY by /generate/ca routes
TEMPLATE_WEEKLY_CA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDKEY_TEMPLATE_WEEKLY.docx')
TEMPLATE_DAILY_CA  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDKEY_TEMPLATE_DAILY.docx')
FORM            = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fundgate_form.html')
DEALS_PAGE      = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'deals_page.html')

SIGNER2_BLOCKS_DIR = os.path.dirname(os.path.abspath(__file__))

FIELDS = {
    '«Agreement_Date»':             'Agreement_Date',
    '«Agreement_Date1»':            'Agreement_Date',
    '«Merchant_Legal_Name»':        'Merchant_Legal_Name',
    '«Merchant_Display_Name»':      'Merchant_Display_Name',
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

    # Final safety net: ensure the merged document.xml has no XML-illegal
    # control characters (covers both contract and disclosure-inserted text).
    contract_xml = strip_illegal_xml_chars(contract_xml)
    contract_files['word/document.xml'] = contract_xml.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data_bytes in contract_files.items():
            zout.writestr(name, data_bytes)
    return buf.getvalue()


def fill_docx(data):
    is_daily = data.get('dealType') == 'daily'
    # Fundkey flag — when True, use Fundkey-branded templates instead of
    # FundGate. Legacy 'isCA' is accepted as an alias for backwards compat.
    is_fundkey = bool(data.get('isFundkey', False) or data.get('isCA', False))
    if is_fundkey:
        template = TEMPLATE_DAILY_CA if is_daily else TEMPLATE_WEEKLY_CA
    else:
        template = TEMPLATE_DAILY if is_daily else TEMPLATE_WEEKLY
    with zipfile.ZipFile(template, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    doc = files['word/document.xml'].decode('utf-8')
    two_signers = data.get('twoSigners', False)

    def safe(val):
        return strip_illegal_xml_chars(val or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

    # On a 2-signer deal, the addendum intro line must name BOTH guarantors.
    # Target only the intro: the «Owner_Guarantor_1» placeholder immediately
    # before the word "Guarantors" (which appears only in the addendum intro).
    # All other placeholder spots are left to the normal field replacement
    # below, so single-signer output is unchanged. Fails safe (no edit) if the
    # expected text isn't found.
    if two_signers and (data.get('Owner_Guarantor_2') or '').strip():
        _g1 = safe((data.get('Owner_Guarantor_1') or '').upper())
        _g2 = safe((data.get('Owner_Guarantor_2') or '').upper())
        _gpos = doc.find('Guarantors')
        if _gpos != -1:
            _ppos = doc.rfind('<w:t>«Owner_Guarantor_1»</w:t>', 0, _gpos)
            if _ppos != -1:
                _rs = doc.rfind('<w:r ', 0, _ppos)
                _re = doc.find('</w:r>', _ppos) + len('</w:r>')
                _run = doc[_rs:_re]
                _both = (_run.replace('«Owner_Guarantor_1»', _g1)
                         + '<w:r><w:rPr><w:color w:val="010202"/></w:rPr>'
                           '<w:t xml:space="preserve"> and </w:t></w:r>'
                         + _run.replace('«Owner_Guarantor_1»', _g2))
                doc = doc[:_rs] + _both + doc[_re:]

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
                # Rule: every signer-2 field prints UPPERCASE on the contract.
                val = (data.get(key, '') or '').upper()
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
        '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Phone: 631-772-9020</w:t></w:r></w:p>'
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

    # ── Fundkey-only: swap the FundGate email line in the consumer notice ────
    # Fundkey email = admin@fundkeyllc.com
    if is_fundkey:
        doc = doc.replace(
            '<w:t>Email: admin@fundgatellc.com</w:t>',
            '<w:t>Email: admin@fundkeyllc.com</w:t>'
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
    # Rule: EVERY field prints UPPERCASE on the contract, EXCEPT the fields in
    # NO_UPPER (kept as entered). ACH_Frequency reads "Weekly"/"Daily" in
    # running text on the ACH debit page, so it stays in regular case.
    # Uppercasing is a no-op on numeric/money/percent/date-slash values, so for
    # everything else this only changes the text fields (names, addresses,
    # entity type, DL, etc.) and guarantees no field is ever missed.
    NO_UPPER = {'ACH_Frequency'}
    # Compute the merchant display name used in select template spots
    # (main signature, ACH top, ACH sig, addendum body, addendum sig).
    # When DBA differs from legal name, append " DBA <name>"; otherwise just
    # use the legal name. Comparison is case-insensitive, ignoring whitespace.
    _legal = (data.get('Merchant_Legal_Name', '') or '').strip()
    _dba   = (data.get('Merchant_DBA', '') or '').strip()
    if _dba and _dba.upper() != _legal.upper():
        data['Merchant_Display_Name'] = f'{_legal} DBA {_dba}'
    else:
        data['Merchant_Display_Name'] = _legal

    for field, key in FIELDS.items():
        raw = data.get(key, '') or ''
        val = raw if key in NO_UPPER else raw.upper()
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
    """Legacy alias retained for any internal callers — same as Fundkey filename."""
    return safe_filename_fundkey(data, ext)

def safe_filename_fundkey(data, ext):
    """Fundkey entity filename. State included so Jack can scan by state."""
    deal = data.get('dealType', 'weekly').capitalize()
    state = (data.get('State_of_Organization') or '').upper().strip()
    state_part = f'_{state}' if state else ''
    dba  = re.sub(r'\s+', '_', data.get('Merchant_DBA') or data.get('Merchant_Legal_Name','Contract'))
    date = (data.get('Agreement_Date','') or '').replace('/','_')
    return f"Fundkey_{deal}{state_part}_{dba}_{date}.{ext}"

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args): pass

    def do_GET(self):
        # Strip query string for path matching (e.g. /?clone=1 -> /)
        from urllib.parse import urlparse
        path_only = urlparse(self.path).path
        if path_only == '/':
            self.send_response(200)
            self.send_header('Content-Type','text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(open(FORM,'rb').read())
        elif path_only == '/deals' or path_only == '/deals/':
            # Past Deals page (HTML always served; client-side calls /deals/check-auth)
            self.send_response(200)
            self.send_header('Content-Type','text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(open(DEALS_PAGE,'rb').read())
        elif self.path == '/deals/check-auth':
            cookie = self.headers.get('Cookie', '')
            if auth_module.is_valid_cookie(cookie):
                self.send_response(200); self.send_header('Content-Type','application/json'); self.end_headers()
                self.wfile.write(b'{"ok":true}')
            else:
                self.send_response(401); self.send_header('Content-Type','application/json'); self.end_headers()
                self.wfile.write(b'{"ok":false}')
        elif self.path.startswith('/deals/search'):
            if not self._require_auth(): return
            from urllib.parse import urlparse, parse_qs
            qs = parse_qs(urlparse(self.path).query)
            q = (qs.get('q', [''])[0] or '').strip()
            entity = (qs.get('entity', [''])[0] or '').strip() or None
            if not db_module.is_configured():
                self.send_response(503); self.send_header('Content-Type','application/json'); self.end_headers()
                self.wfile.write(b'{"error":"db not configured"}')
                return
            results = db_module.search_deals(query=q, entity=entity)
            self.send_response(200); self.send_header('Content-Type','application/json'); self.end_headers()
            self.wfile.write(json.dumps(results, default=str).encode())
        elif self.path.startswith('/deals/') and len(self.path.split('/')) == 3:
            # GET /deals/<id>
            if not self._require_auth(): return
            deal_id = self.path.split('/')[2]
            deal = db_module.get_deal(deal_id)
            if not deal:
                self.send_response(404); self.send_header('Content-Type','application/json'); self.end_headers()
                self.wfile.write(b'{"error":"not found"}')
                return
            self.send_response(200); self.send_header('Content-Type','application/json'); self.end_headers()
            self.wfile.write(json.dumps(deal, default=str).encode())
        else:
            self.send_response(404); self.end_headers()

    def _require_auth(self):
        """Check session cookie; send 401 and return False if invalid. Else True."""
        cookie = self.headers.get('Cookie', '')
        if auth_module.is_valid_cookie(cookie):
            return True
        self.send_response(401); self.send_header('Content-Type','application/json'); self.end_headers()
        self.wfile.write(b'{"error":"unauthorized"}')
        return False

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
                # NY, CT and UT each use their own dedicated disclosure module (FundGate-only).
                # All other states continue through the existing builder.
                state = (data.get('State_of_Organization') or '').upper().strip()
                if state == 'NY':
                    disc_bytes = build_ny_disclosure_bytes(data)
                elif state == 'CT':
                    disc_bytes = build_ct_disclosure_bytes(data)
                elif state == 'UT':
                    disc_bytes = build_ut_disclosure_bytes(data)
                else:
                    disc_bytes = build_disclosure_bytes(data)
                # No-state Commercial Financing Disclosure is used ONLY when the
                # deal's state has no CFDL disclosure of its own; it never
                # replaces a state disclosure. The Prior_Balance_Amount field
                # flows into whichever disclosure renders.
                if disc_bytes is None and data.get('includeNoStateCFD'):
                    disc_bytes = build_nostate_disclosure_bytes(data)
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
                # Best-effort auto-save to deals DB (never blocks the response).
                try:
                    deal_type = (data.get('dealType') or '').lower() or 'weekly'
                    db_module.save_deal(data, entity='FundGate', deal_type=deal_type)
                except Exception as _save_err:
                    print(f'[db] auto-save failed: {_save_err}')
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type','application/json')
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        elif self.path in ('/generate/fundkey', '/generate/fundkey/pdf',
                            '/generate/ca', '/generate/ca/pdf'):
            # Fundkey-route — uses Fundkey-branded templates. State of
            # Organization is honored from the form (CA, FL, GA, etc.).
            # Disclosure rendering:
            #   - state = CA            -> CA disclosure (Fundkey-branded)
            #   - state in FL/GA/LA/... -> state disclosure (Fundkey-branded)
            #   - other states          -> no disclosure (same as FundGate)
            # The legacy /generate/ca paths are kept as aliases so existing
            # bookmarks still work.
            want_pdf = self.path.endswith('/pdf')
            length = int(self.headers.get('Content-Length', 0))
            data = json.loads(self.rfile.read(length))
            try:
                # Flag the request as a Fundkey deal — drives template
                # selection in fill_docx AND provider branding in
                # build_disclosure_bytes.
                data['isFundkey'] = True
                data['isCA'] = True  # legacy alias used by fill_docx
                docx_bytes = fill_docx(data)
                # Pick disclosure by state
                state = (data.get('State_of_Organization') or '').upper().strip()
                if state == 'CA':
                    disc_bytes = build_ca_disclosure_bytes(data)
                else:
                    disc_bytes = build_disclosure_bytes(data)
                # No-state CFD only when the state has no disclosure of its own.
                if disc_bytes is None and data.get('includeNoStateCFD'):
                    disc_bytes = build_nostate_disclosure_bytes(data)
                if disc_bytes:
                    docx_bytes = merge_disclosure_into_contract(docx_bytes, disc_bytes)
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime      = 'application/pdf'
                    fname     = safe_filename_fundkey(data, 'pdf')
                else:
                    out_bytes = docx_bytes
                    mime      = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname     = safe_filename_fundkey(data, 'docx')
                self.send_response(200)
                self.send_header('Content-Type', mime)
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"; filename*=UTF-8''{fname}')
                self.send_header('Content-Length', str(len(out_bytes)))
                self.send_header('Access-Control-Allow-Origin','*')
                self.end_headers()
                self.wfile.write(out_bytes)
                # Best-effort auto-save to deals DB (never blocks the response).
                try:
                    deal_type = (data.get('dealType') or '').lower() or 'weekly'
                    db_module.save_deal(data, entity='Fundkey', deal_type=deal_type)
                except Exception as _save_err:
                    print(f'[db] auto-save failed: {_save_err}')
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
                prefix = 'FundKey' if (data.get('entity') or '').strip().lower() == 'fundkey' else 'FundGate'
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime = 'application/pdf'
                    fname = f'{prefix}_Payoff_{merchant}_{date_str}.pdf'
                else:
                    out_bytes = docx_bytes
                    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname = f'{prefix}_Payoff_{merchant}_{date_str}.docx'
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
                prefix = 'FundKey' if (data.get('entity') or '').strip().lower() == 'fundkey' else 'FundGate'
                if want_pdf:
                    out_bytes = docx_to_pdf(docx_bytes)
                    mime = 'application/pdf'
                    fname = f'{prefix}_ZeroBalance_{merchant}_{date_str}.pdf'
                else:
                    out_bytes = docx_bytes
                    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    fname = f'{prefix}_ZeroBalance_{merchant}_{date_str}.docx'
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
        elif self.path == '/deals/login':
            length = int(self.headers.get('Content-Length', 0))
            try:
                payload = json.loads(self.rfile.read(length))
            except Exception:
                payload = {}
            if auth_module.check_password(payload.get('password')):
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Set-Cookie', auth_module.cookie_header())
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
            else:
                self.send_response(401)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"ok":false}')
        elif self.path == '/deals/logout':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Set-Cookie', auth_module.clear_cookie_header())
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        else:
            self.send_response(404); self.end_headers()

    def do_DELETE(self):
        if self.path.startswith('/deals/') and len(self.path.split('/')) == 3:
            if not self._require_auth(): return
            deal_id = self.path.split('/')[2]
            ok = db_module.delete_deal(deal_id)
            if ok:
                self.send_response(204); self.end_headers()
            else:
                self.send_response(500); self.send_header('Content-Type', 'application/json'); self.end_headers()
                self.wfile.write(b'{"error":"delete failed"}')
        else:
            self.send_response(404); self.end_headers()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    server = HTTPServer(('0.0.0.0', port), Handler)
    print(f'FundGate server running on port {port}')
    server.serve_forever()
