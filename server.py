import os
import sys
import json
import subprocess
import platform
import socketserver
import time
import mimetypes
import urllib.parse
from http.server import SimpleHTTPRequestHandler

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 5500
HOST = "127.0.0.1"

def bring_window_to_front(target_name):
    """Kích hoạt và đưa cửa sổ vừa mở lên Foreground trên Windows"""
    if platform.system() != "Windows" or not target_name:
        return
    try:
        import ctypes
        ctypes.windll.user32.AllowSetForegroundWindow(-1)
    except Exception:
        pass

    try:
        clean_name = target_name.replace("'", "''").replace('"', '')
        ps_cmd = f"""
        Start-Sleep -Milliseconds 300
        $ws = New-Object -ComObject Wscript.Shell
        $ws.AppActivate('{clean_name}')
        """
        subprocess.Popen(['powershell.exe', '-NoProfile', '-WindowStyle', 'Hidden', '-Command', ps_cmd], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

class MAIORequestHandler(SimpleHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        # Endpoint xem / tải trực tiếp tệp tin nội bộ từ ổ cứng máy tính
        if self.path.startswith('/api/view-file'):
            try:
                parsed = urllib.parse.urlparse(self.path)
                qs = urllib.parse.parse_qs(parsed.query)
                target_path = qs.get('path', [''])[0].strip()
                if not target_path or not os.path.exists(target_path):
                    self.send_error(404, "File not found on local disk")
                    return

                clean_path = os.path.normpath(target_path)
                if os.path.isdir(clean_path):
                    self.send_error(400, "Target is a folder, not a file")
                    return

                ctype, _ = mimetypes.guess_type(clean_path)
                if not ctype:
                    low = clean_path.lower()
                    if low.endswith('.pdf'):
                        ctype = 'application/pdf'
                    elif low.endswith('.xlsx'):
                        ctype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    elif low.endswith('.xls'):
                        ctype = 'application/vnd.ms-excel'
                    elif low.endswith('.docx'):
                        ctype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    elif low.endswith('.doc'):
                        ctype = 'application/msword'
                    else:
                        ctype = 'application/octet-stream'

                filename = os.path.basename(clean_path)
                encoded_filename = urllib.parse.quote(filename)

                with open(clean_path, 'rb') as f:
                    content = f.read()

                self.send_response(200)
                self.send_header('Content-Type', ctype)
                self.send_header('Content-Length', str(len(content)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Disposition', f"inline; filename*=UTF-8''{encoded_filename}")
                self.end_headers()
                self.wfile.write(content)
                return
            except Exception as e:
                self.send_error(500, f"Error reading file: {e}")
                return

        super().do_GET()

    def do_POST(self):
        if self.path == '/api/open-local':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length).decode('utf-8')
                data = json.loads(post_data) if post_data else {}
                target_path = data.get('path', '').strip()
                is_folder = data.get('is_folder', False)
                action = data.get('action', 'open') # 'open' hoặc 'select'

                if not target_path:
                    self._send_json(400, {"success": False, "error": "Path is empty"})
                    return

                clean_path = os.path.normpath(target_path)

                if platform.system() == "Windows":
                    try:
                        import ctypes
                        ctypes.windll.user32.AllowSetForegroundWindow(-1)
                    except Exception:
                        pass

                    # 1. Mở thư mục
                    if is_folder or os.path.isdir(clean_path):
                        folder_to_open = clean_path if os.path.isdir(clean_path) else os.path.dirname(clean_path)
                        p = folder_to_open
                        while p and not os.path.exists(p):
                            np = os.path.dirname(p)
                            if np == p:
                                break
                            p = np
                        final_folder = p if (p and os.path.exists(p)) else folder_to_open

                        # Mở Explorer
                        subprocess.Popen(['explorer.exe', final_folder])
                        # Kích hoạt cửa sổ Explorer lên trên cùng
                        bring_window_to_front(os.path.basename(final_folder))

                        self._send_json(200, {
                            "success": True,
                            "type": "folder",
                            "path": final_folder,
                            "name": os.path.basename(final_folder)
                        })
                        return

                    # 2. Chọn file trong thư mục
                    if action == 'select' and os.path.isfile(clean_path):
                        subprocess.Popen(['explorer.exe', f'/select,{clean_path}'])
                        bring_window_to_front(os.path.basename(os.path.dirname(clean_path)))
                        self._send_json(200, {
                            "success": True,
                            "type": "file_select",
                            "path": clean_path,
                            "name": os.path.basename(clean_path)
                        })
                        return

                    # 3. Mở tệp tin
                    if os.path.isfile(clean_path):
                        try:
                            os.startfile(clean_path)
                        except Exception:
                            subprocess.Popen(f'cmd.exe /c start "" "{clean_path}"', shell=True)
                        
                        bring_window_to_front(os.path.basename(clean_path))
                        self._send_json(200, {
                            "success": True,
                            "type": "file",
                            "path": clean_path,
                            "name": os.path.basename(clean_path),
                            "view_url": f"/api/view-file?path={urllib.parse.quote(clean_path)}"
                        })
                        return

                    # Nếu tệp tin không tồn tại, mở thư mục cha gần nhất
                    p = clean_path
                    while p and not os.path.exists(p):
                        np = os.path.dirname(p)
                        if np == p:
                            break
                        p = np
                    if p and os.path.exists(p):
                        subprocess.Popen(['explorer.exe', p])
                        bring_window_to_front(os.path.basename(p))
                        self._send_json(200, {
                            "success": True,
                            "type": "fallback_folder",
                            "path": p,
                            "name": os.path.basename(p),
                            "warning": "Tệp tin chưa tồn tại, đã mở thư mục lưu trữ!"
                        })
                    else:
                        subprocess.Popen(['explorer.exe', clean_path])
                        self._send_json(200, {"success": True, "path": clean_path})

                elif platform.system() == "Darwin":
                    if is_folder or os.path.isdir(clean_path):
                        subprocess.Popen(['open', clean_path if os.path.isdir(clean_path) else os.path.dirname(clean_path)])
                    else:
                        subprocess.Popen(['open', clean_path])
                    self._send_json(200, {"success": True, "path": clean_path})
                else:
                    if is_folder or os.path.isdir(clean_path):
                        subprocess.Popen(['xdg-open', clean_path if os.path.isdir(clean_path) else os.path.dirname(clean_path)])
                    else:
                        subprocess.Popen(['xdg-open', clean_path])
                    self._send_json(200, {"success": True, "path": clean_path})

            except Exception as e:
                self._send_json(500, {"success": False, "error": str(e)})
            return

        self.send_error(404, "Endpoint not found")

    def _send_json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

if __name__ == '__main__':
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer((HOST, PORT), MAIORequestHandler) as httpd:
        print(f"MAIO Server running at http://{HOST}:{PORT}/", flush=True)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.", flush=True)
