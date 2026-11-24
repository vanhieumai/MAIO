import os
import sys
import json
import subprocess
import platform
import socketserver
from http.server import SimpleHTTPRequestHandler

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 5500
HOST = "127.0.0.1"

class MAIORequestHandler(SimpleHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        if self.path == '/api/open-local':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length).decode('utf-8')
                data = json.loads(post_data) if post_data else {}
                target_path = data.get('path', '').strip()
                if not target_path:
                    self._send_json(400, {"success": False, "error": "Path is empty"})
                    return

                clean_path = os.path.normpath(target_path)
                
                # Mở trên hệ điều hành tương ứng
                if platform.system() == "Windows":
                    if os.path.exists(clean_path):
                        os.startfile(clean_path)
                    else:
                        # Mở thư mục cha gần nhất nếu file/thư mục con chưa tạo
                        p = clean_path
                        while p and not os.path.exists(p):
                            np = os.path.dirname(p)
                            if np == p:
                                break
                            p = np
                        if p and os.path.exists(p):
                            os.startfile(p)
                        else:
                            subprocess.Popen(['explorer', clean_path])
                elif platform.system() == "Darwin":
                    subprocess.Popen(['open', clean_path])
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
        print(f"MAIO Server running at http://{HOST}:{PORT}/")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
