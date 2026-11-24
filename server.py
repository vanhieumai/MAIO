import os
import sys
import json
import subprocess
import platform
import socketserver
import time
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
        # Sử dụng PowerShell Wscript.Shell để active cửa sổ có tiêu đề chứa target_name
        clean_name = target_name.replace("'", "''").replace('"', '')
        ps_cmd = f"""
        Start-Sleep -Milliseconds 400
        $ws = New-Object -ComObject Wscript.Shell
        $ws.AppActivate('{clean_name}')
        """
        subprocess.Popen(['powershell.exe', '-NoProfile', '-WindowStyle', 'Hidden', '-Command', ps_cmd])
    except Exception:
        pass

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
                is_folder = data.get('is_folder', False)
                action = data.get('action', 'open') # 'open' hoặc 'select'

                if not target_path:
                    self._send_json(400, {"success": False, "error": "Path is empty"})
                    return

                clean_path = os.path.normpath(target_path)
                print(f"[API open-local] Received: {clean_path} (is_folder={is_folder}, action={action})")

                if platform.system() == "Windows":
                    # Đảm bảo Windows cho phép tiến trình mới chiếm foreground
                    try:
                        import ctypes
                        ctypes.windll.user32.AllowSetForegroundWindow(-1)
                    except Exception:
                        pass

                    # 1. Trường hợp yêu cầu mở thư mục hoặc đường dẫn thực tế là thư mục
                    if is_folder or os.path.isdir(clean_path):
                        folder_to_open = clean_path if os.path.isdir(clean_path) else os.path.dirname(clean_path)
                        # Tìm thư mục cha gần nhất nếu thư mục con chưa tồn tại
                        p = folder_to_open
                        while p and not os.path.exists(p):
                            np = os.path.dirname(p)
                            if np == p:
                                break
                            p = np
                        final_folder = p if (p and os.path.exists(p)) else folder_to_open

                        # Mở Explorer trực tiếp vào thư mục
                        subprocess.Popen(['explorer.exe', final_folder])
                        bring_window_to_front(os.path.basename(final_folder))
                        self._send_json(200, {
                            "success": True,
                            "type": "folder",
                            "path": final_folder,
                            "name": os.path.basename(final_folder)
                        })
                        return

                    # 2. Trường hợp yêu cầu chọn (highlight) file trong thư mục
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

                    # 3. Trường hợp mở tệp tin
                    if os.path.isfile(clean_path):
                        os.startfile(clean_path)
                        bring_window_to_front(os.path.basename(clean_path))
                        self._send_json(200, {
                            "success": True,
                            "type": "file",
                            "path": clean_path,
                            "name": os.path.basename(clean_path)
                        })
                        return

                    # Nếu tệp tin không tồn tại, mở thư mục cha gần nhất và báo
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
                print(f"[API open-local ERROR] {e}")
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
