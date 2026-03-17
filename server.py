#!/usr/bin/env python3
"""FundGate Weekly Contract Generator — Render.com deployment"""
import http.server, json, zipfile, re, os, subprocess, tempfile, shutil
from http.server import HTTPServer, BaseHTTPRequestHandler

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FUNDGATE_TEMPLATE_WEEKLY.docx')
FORM     = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fundgate_form.html')

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

def fill_docx(data):
    with zipfile.ZipFile(TEMPLATE, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}
    doc = files['word/document.xml'].decode('utf-8')
    for field, key in FIELDS.items():
        val = data.get(key, '')
        safe = val.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        doc = doc.replace(field, safe)
    files['word/document.xml'] = doc.encode('utf-8')
    import io
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
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, docx_path],
            capture_output=True, timeout=60
        )
        if os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                return f.read()
        raise Exception('PDF conversion failed')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def safe_filename(data, ext):
    dba  = re.sub(r'\s+', '_', data.get('Merchant_DBA') or data.get('Merchant_Legal_Name','Contract'))
    date = (data.get('Agreement_Date','') or '').replace('/','_')
    return f"FundGate_Weekly_{dba}_{date}.{ext}"

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
