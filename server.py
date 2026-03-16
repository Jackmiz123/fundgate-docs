#!/usr/bin/env python3
"""FundGate Weekly Contract Generator — Render.com deployment"""
import http.server, json, zipfile, re, os, subprocess, tempfile, shutil, io
from http.server import HTTPServer, BaseHTTPRequestHandler
from disclosure_module import build_disclosure_bytes

TEMPLATE_WEEKLY = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_WEEKLY.docx')
TEMPLATE_DAILY  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_DAILY.docx')
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
    '«Repurchase_31_60_Day_Amount»':'Repurchase_31_60_Day_Amount',
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
    template = TEMPLATE_DAILY if data.get('dealType') == 'daily' else TEMPLATE_WEEKLY
    with zipfile.ZipFile(template, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    doc = files['word/document.xml'].decode('utf-8')
    two_signers = data.get('twoSigners', False)

    def safe(val):
        return (val or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

    # ── Handle signer 2 placeholder blocks ──────────────────────────────────────
    if two_signers:
        def fill_s2_block(block_xml):
            for placeholder, key in FIELDS_S2.items():
                val = data.get(key, '') or ''
                if key in ('Owner_Guarantor_2', 'Merchant_Legal_Name', 'Merchant_DBA'):
                    val = val.upper()
                block_xml = block_xml.replace(placeholder, safe(val))
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
    doc = doc.replace('«SIGNER2_BLOCK_ACH»',        block_ach)
    doc = doc.replace('«SIGNER2_BLOCK_BANKLOGIN»',  block_bank)
    doc = doc.replace('«SIGNER2_BLOCK_ADDENDUM»',   block_add)
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
                self.send_header('Content-Disposition', f'attachment; filename="{fname}"')
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
        else:
            self.send_response(404); self.end_headers()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    server = HTTPServer(('0.0.0.0', port), Handler)
    print(f'FundGate server running on port {port}')
    server.serve_forever()
