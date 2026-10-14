#!/usr/bin/env bash
#!/usr/bin/env python3
"""
Enterprise Retail Price Tag Generator Server
Lightweight, zero-dependency Python 3 server providing:
- High-speed static file serving
- Robust Excel (.xlsx) extraction via built-in zipfile/xml
- Google Sheets proxy to bypass browser CORS restrictions
"""

import http.server
import socketserver
import urllib.parse
import urllib.request
import os
import sys
import json
import io
import re
import zipfile
import xml.etree.ElementTree as ET
import html

PORT = int(os.environ.get('PORT', 8080))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def col_letter_to_index(col_str):
    """Converts Excel column letters like 'A' -> 0, 'B' -> 1, 'AA' -> 26."""
    idx = 0
    for char in col_str.upper():
        if 'A' <= char <= 'Z':
            idx = idx * 26 + (ord(char) - ord('A') + 1)
    return max(0, idx - 1)

def parse_cell_ref(cell_ref):
    """Splits 'C2' into column index and row index."""
    match = re.match(r'([A-Za-z]+)([0-9]+)', cell_ref)
    if match:
        col_letters, row_num = match.groups()
        return col_letter_to_index(col_letters), int(row_num)
    return None, None

def parse_xlsx(file_bytes):
    """Parses .xlsx file into 2D array of strings using Python standard library."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(file_bytes))
        shared_strings = []
        if 'xl/sharedStrings.xml' in zf.namelist():
            tree = ET.fromstring(zf.read('xl/sharedStrings.xml'))
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in tree.findall('.//ns:si', ns):
                # An si can have direct <t> or multiple <r><t> rich text runs
                text_parts = [t.text for t in si.findall('.//ns:t', ns) if t.text]
                shared_strings.append(''.join(text_parts))
        
        # Find first sheet from namelist
        sheet_names = [name for name in zf.namelist() if name.startswith('xl/worksheets/sheet') and name.endswith('.xml')]
        sheet_names.sort()
        target_sheet = sheet_names[0] if sheet_names else 'xl/worksheets/sheet1.xml'
        
        if target_sheet not in zf.namelist():
            return []
            
        tree = ET.fromstring(zf.read(target_sheet))
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        rows_data = []
        for row in tree.findall('.//ns:row', ns):
            cells = row.findall('./ns:c', ns)
            if not cells:
                continue

            max_col = 0
            row_map = {}

            for cell in cells:
                ref = cell.get('r', '')
                col_idx, _ = parse_cell_ref(ref)
                if col_idx is None:
                    col_idx = len(row_map)

                max_col = max(max_col, col_idx)

                cell_type = cell.get('t')
                val_el = cell.find('./ns:v', ns)
                inline_text = cell.find('.//ns:t', ns)
                
                if inline_text is not None and inline_text.text:
                    val = inline_text.text
                elif val_el is not None and val_el.text:
                    val = val_el.text
                    if cell_type == 's' and val.isdigit():
                        idx = int(val)
                        val = shared_strings[idx] if idx < len(shared_strings) else val
                    elif cell_type == 'b':
                        val = 'TRUE' if val == '1' else 'FALSE'
                else:
                    val = ''
                
                row_map[col_idx] = val.strip()

            # Construct row array preserving empty columns
            curr_row = [row_map.get(c, '') for c in range(max_col + 1)]
            if any(curr_row):
                rows_data.append(curr_row)
                
        return rows_data
    except Exception as e:
        print(f"Error parsing xlsx: {e}", file=sys.stderr)
        return []

class PriceTagRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/api/proxy-image':
            params = urllib.parse.parse_qs(parsed.query)
            target_url = params.get('url', [''])[0]
            file_id = params.get('id', [''])[0]
            if not target_url and file_id:
                target_url = f"https://lh3.googleusercontent.com/d/{file_id}"

            if not target_url:
                self.send_response(400)
                self.end_headers()
                return

            try:
                req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=15) as resp:
                    c_type = resp.headers.get('Content-Type', 'image/jpeg')
                    img_data = resp.read()
                    self.send_response(200)
                    self.send_header('Content-Type', c_type)
                    self.send_header('Cache-Control', 'public, max-age=86400')
                    self.end_headers()
                    self.wfile.write(img_data)
                    return
            except Exception as e:
                self.send_response(502)
                self.end_headers()
                return

        super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/parse-excel':
            content_length = int(self.headers.get('Content-Length', 0))
            raw_data = self.rfile.read(content_length)
            
            content_type = self.headers.get('Content-Type', '')
            file_bytes = raw_data

            # If multipart/form-data was sent, extract boundary accurately
            if 'multipart/form-data' in content_type:
                # Find boundary marker
                boundary_match = re.search(r'boundary=([^;]+)', content_type)
                if boundary_match:
                    boundary = boundary_match.group(1).strip('"\'').encode('utf-8')
                    delimiter = b'--' + boundary
                    parts = raw_data.split(delimiter)
                    for part in parts:
                        if b'\r\n\r\n' in part:
                            headers_part, body_part = part.split(b'\r\n\r\n', 1)
                            # Remove trailing \r\n if present
                            if body_part.endswith(b'\r\n'):
                                body_part = body_part[:-2]
                            # Check if valid zip
                            if body_part.startswith(b'PK\x03\x04'):
                                file_bytes = body_part
                                break

            rows = parse_xlsx(file_bytes)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': bool(rows), 'rows': rows}).encode('utf-8'))
            return

        elif path == '/api/fetch-gsheet':
            content_length = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_length)
            try:
                data = json.loads(post_body.decode('utf-8'))
                url = data.get('url', '')
                
                match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
                if not match:
                    raise ValueError("Invalid Google Sheets URL")
                
                sheet_id = match.group(1)
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                
                req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=10) as response:
                    csv_text = response.read().decode('utf-8', errors='replace')
                    
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': True, 'csv': csv_text}).encode('utf-8'))
                return
            except Exception as e:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode('utf-8'))
                return

        elif path == '/api/fetch-gdrive-folder':
            content_length = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_length)
            try:
                data = json.loads(post_body.decode('utf-8'))
                folder_input = data.get('url', '').strip() or data.get('folderId', '').strip()
                api_key = data.get('apiKey', '').strip()

                folder_id = None
                m1 = re.search(r'folders/([a-zA-Z0-9_-]+)', folder_input)
                if m1:
                    folder_id = m1.group(1)
                else:
                    m2 = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', folder_input)
                    if m2:
                        folder_id = m2.group(1)
                    elif re.match(r'^[a-zA-Z0-9_-]{15,60}$', folder_input):
                        folder_id = folder_input

                if not folder_id:
                    raise ValueError("Could not extract a valid Google Drive Folder ID. Please paste a full Google Drive folder link or folder ID.")

                files_found = {}
                image_exts = r'\.(?:jpg|jpeg|png|webp|gif|avif|bmp|heic)'

                # 1. If API key provided, query official Google Drive API v3
                if api_key:
                    try:
                        q_query = urllib.parse.quote(f"'{folder_id}' in parents and trashed=false")
                        api_url = f"https://www.googleapis.com/drive/v3/files?q={q_query}&fields=files(id,name,mimeType)&pageSize=1000&key={api_key}"
                        req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=12) as response:
                            api_res = json.loads(response.read().decode('utf-8'))
                            for f in api_res.get('files', []):
                                mime = f.get('mimeType', '')
                                name = f.get('name', '')
                                fid = f.get('id', '')
                                if mime.startswith('image/') or re.search(image_exts + r'$', name, re.I):
                                    files_found[fid] = name
                    except Exception as err:
                        print("Google Drive API fetch warning:", err)

                # 2. Public folder scrape via embeddedfolderview and folders view
                headers = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
                urls_to_try = [
                    f"https://drive.google.com/embeddedfolderview?id={folder_id}",
                    f"https://drive.google.com/drive/folders/{folder_id}"
                ]

                patterns = [
                    r'\[\"([a-zA-Z0-9_-]{20,})\",\[\"([^\"\\\\]+' + image_exts + r')\"',
                    r'\[\"([a-zA-Z0-9_-]{20,})\",\s*\"([^\"\\\\]+' + image_exts + r')\"',
                    r'data-id=[\"\']([a-zA-Z0-9_-]{20,})[\"\'][^>]*aria-label=[\"\']([^\"\']+' + image_exts + r')[\"\']',
                    r'aria-label=[\"\']([^\"\']+' + image_exts + r')[\"\'][^>]*data-id=[\"\']([a-zA-Z0-9_-]{20,})[\"\']',
                    r'/file/d/([a-zA-Z0-9_-]{20,})[^\"]*\"[^>]*>([^<]+' + image_exts + r')<',
                    r'href=[\"\'][^\"\']*?/file/d/([a-zA-Z0-9_-]{20,})[^\"\']*[\"\'][^>]*>([^<]+?' + image_exts + r')<',
                    r'entry-([a-zA-Z0-9_-]{20,})[\s\S]*?<a[^>]*>([^<]+' + image_exts + r')<',
                    r'\"id\"\s*:\s*\"([a-zA-Z0-9_-]{20,})\"\s*,\s*\"title\"\s*:\s*\"([^\"\']+' + image_exts + r')\"',
                    r'\"title\"\s*:\s*\"([^\"\']+' + image_exts + r')\"\s*,\s*\"id\"\s*:\s*\"([a-zA-Z0-9_-]{20,})\"',
                    r'\"name\"\s*:\s*\"([^\"\']+' + image_exts + r')\"[\s\S]*?\"id\"\s*:\s*\"([a-zA-Z0-9_-]{20,})\"',
                    r'\"id\"\s*:\s*\"([a-zA-Z0-9_-]{20,})\"[\s\S]*?\"name\"\s*:\s*\"([^\"\']+' + image_exts + r')\"',
                ]

                for u in urls_to_try:
                    try:
                        req = urllib.request.Request(u, headers=headers)
                        with urllib.request.urlopen(req, timeout=12) as response:
                            html_text = response.read().decode('utf-8', errors='replace')
                            for pat in patterns:
                                matches = re.findall(pat, html_text, re.IGNORECASE)
                                for match in matches:
                                    if len(match) == 2:
                                        p1, p2 = match[0].strip(), match[1].strip()
                                        clean_p1 = html.unescape(p1).replace('\\', '')
                                        clean_p2 = html.unescape(p2).replace('\\', '')
                                        if re.search(image_exts, clean_p2, re.I) and not re.search(image_exts, clean_p1, re.I):
                                            files_found[clean_p1] = clean_p2
                                        elif re.search(image_exts, clean_p1, re.I) and not re.search(image_exts, clean_p2, re.I):
                                            files_found[clean_p2] = clean_p1
                    except Exception as e_scrape:
                        print(f"Error scraping {u}:", e_scrape)

                file_list = [
                    {
                        'id': fid,
                        'name': fname,
                        'url': f"https://lh3.googleusercontent.com/d/{fid}"
                    }
                    for fid, fname in files_found.items()
                ]

                if not file_list:
                    raise ValueError(
                        "No images found in this Google Drive folder. "
                        "Please ensure the folder is shared as 'Anyone with the link can view'. "
                        "If this is a corporate or restricted folder, enter an API key or use the local folder detector."
                    )

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': True, 'folderId': folder_id, 'files': file_list}).encode('utf-8'))
                return
            except Exception as e:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode('utf-8'))
                return

        self.send_response(404)
        self.end_headers()

def run(port=PORT):
    current_port = port
    while current_port < port + 20:
        try:
            with socketserver.TCPServer(("", current_port), PriceTagRequestHandler) as httpd:
                print(f"===============================================================")
                print(f"🚀 Retail Price Tag Generator Server running at:")
                print(f"👉  http://localhost:{current_port}")
                print(f"===============================================================")
                sys.stdout.flush()
                httpd.serve_forever()
                break
        except OSError:
            current_port += 1

if __name__ == '__main__':
    run()
