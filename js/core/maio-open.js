/**
 * MAIO-Open Protocol Helper
 * Đăng ký và gọi giao thức maio-open:// để mở thẳng file/thư mục Windows từ web.
 */

const STORAGE_KEY = 'maio_open_protocol_status'; // 'installed' | 'browser_only' | ''

class MaioOpenHelper {
  isWindowsDesktop() {
    const ua = navigator.userAgent || '';
    const isMobile = /Android|iPhone|iPad|iPod|Mobile/i.test(ua);
    const isWin = /Windows/i.test(ua) || /Win/i.test(navigator.platform || '');
    return isWin && !isMobile;
  }

  getStatus() {
    try {
      return localStorage.getItem(STORAGE_KEY) || '';
    } catch (_) {
      return '';
    }
  }

  setStatus(status) {
    try {
      if (status) {
        localStorage.setItem(STORAGE_KEY, status);
      } else {
        localStorage.removeItem(STORAGE_KEY);
      }
    } catch (_) {}
    this.updateSettingsUI();
  }

  updateSettingsUI() {
    const badge = document.getElementById('maio-open-status-badge');
    const resetBtn = document.getElementById('btn-reset-maio-open');
    if (!badge) return;

    const status = this.getStatus();
    if (status === 'installed') {
      badge.textContent = 'Đã cài đặt (Mở thẳng file & thư mục trên Windows)';
      badge.style.color = 'var(--success)';
      if (resetBtn) resetBtn.classList.remove('hidden');
    } else if (status === 'browser_only') {
      badge.textContent = 'Đang dùng chế độ trình duyệt (Đã tắt nhắc cài)';
      badge.style.color = 'var(--text-muted)';
      if (resetBtn) resetBtn.classList.remove('hidden');
    } else {
      badge.textContent = 'Chưa cài đặt (Sẽ hỏi khi bấm Mở file)';
      badge.style.color = 'var(--warning, #d97706)';
      if (resetBtn) resetBtn.classList.add('hidden');
    }
  }

  buildInstallerBatContent() {
    // Chỉ dùng ASCII thuần để tránh lỗi encoding CMD trên Windows
    const psLines = [
      "param([string]$Uri)",
      "if (-not $Uri) { exit 0 }",
      "$raw = $Uri -replace '^maio-open://', ''",
      "$parts = $raw -split '\\?', 2",
      "$action = ($parts[0]).Trim('/').ToLower()",
      "if ($action -eq 'ping') { exit 0 }",
      "$targetPath = ''",
      "if ($parts.Count -gt 1) {",
      "  foreach ($kv in ($parts[1] -split '&')) {",
      "    if ($kv -match '^path=(.*)$') {",
      "      $targetPath = [System.Uri]::UnescapeDataString(($Matches[1] -replace '\\+', '%20'))",
      "    }",
      "  }",
      "}",
      "$targetPath = $targetPath.Trim()",
      "if (-not $targetPath) { exit 0 }",
      "if ($targetPath -notmatch '^([A-Za-z]:\\\\|\\\\\\\\)') { exit 1 }",
      "$norm = $targetPath -replace '/', '\\\\'",
      "if ($action -eq 'folder') {",
      "  if (Test-Path -LiteralPath $norm -PathType Leaf) {",
      "    Start-Process explorer.exe -ArgumentList ('/select,\"' + $norm + '\"')",
      "    exit 0",
      "  }",
      "  $p = $norm",
      "  while ($p -and -not (Test-Path -LiteralPath $p)) {",
      "    $parent = Split-Path -Path $p -Parent",
      "    if (-not $parent -or $parent -eq $p) { break }",
      "    $p = $parent",
      "  }",
      "  if ($p -and (Test-Path -LiteralPath $p)) {",
      "    Start-Process explorer.exe -ArgumentList ('\"' + $p + '\"')",
      "  }",
      "  exit 0",
      "}",
      "if ($action -eq 'select' -and (Test-Path -LiteralPath $norm -PathType Leaf)) {",
      "  Start-Process explorer.exe -ArgumentList ('/select,\"' + $norm + '\"')",
      "  exit 0",
      "}",
      "if (Test-Path -LiteralPath $norm -PathType Leaf) {",
      "  Start-Process -FilePath $norm",
      "  exit 0",
      "}",
      "if (Test-Path -LiteralPath $norm -PathType Container) {",
      "  Start-Process explorer.exe -ArgumentList ('\"' + $norm + '\"')",
      "  exit 0",
      "}",
      "$p = $norm",
      "while ($p -and -not (Test-Path -LiteralPath $p)) {",
      "  $parent = Split-Path -Path $p -Parent",
      "  if (-not $parent -or $parent -eq $p) { break }",
      "  $p = $parent",
      "}",
      "if ($p -and (Test-Path -LiteralPath $p)) {",
      "  Start-Process explorer.exe -ArgumentList ('\"' + $p + '\"')",
      "}"
    ];

    const vbsLines = [
      'Set objArgs = WScript.Arguments',
      'If objArgs.Count > 0 Then',
      '  uriArg = Replace(objArgs(0), Chr(34), Chr(34) & Chr(34))',
      '  Set sh = CreateObject("WScript.Shell")',
      '  sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""%LOCALAPPDATA%\\MAIO\\maio-open.ps1"" """ & uriArg & """", 0, False',
      'End If'
    ];

    const psBase64 = btoa(unescape(encodeURIComponent(psLines.join('\r\n'))));
    const vbsBase64 = btoa(unescape(encodeURIComponent(vbsLines.join('\r\n'))));

    const bat = [
      "@echo off",
      "title MAIO-Open Installer",
      "set \"INSTALL_DIR=%LOCALAPPDATA%\\MAIO\"",
      "if not exist \"%INSTALL_DIR%\" mkdir \"%INSTALL_DIR%\"",
      "set \"PS_FILE=%INSTALL_DIR%\\maio-open.ps1\"",
      "set \"VBS_FILE=%INSTALL_DIR%\\maio-open-launcher.vbs\"",
      "",
      `powershell -NoProfile -Command "[IO.File]::WriteAllBytes('%PS_FILE%', [Convert]::FromBase64String('${psBase64}'))"`,
      `powershell -NoProfile -Command "[IO.File]::WriteAllBytes('%VBS_FILE%', [Convert]::FromBase64String('${vbsBase64}'))"`,
      "",
      "reg add \"HKCU\\Software\\Classes\\maio-open\" /ve /d \"URL:MAIO Open Protocol\" /f >nul",
      "reg add \"HKCU\\Software\\Classes\\maio-open\" /v \"URL Protocol\" /d \"\" /f >nul",
      "reg add \"HKCU\\Software\\Classes\\maio-open\\DefaultIcon\" /ve /d \"explorer.exe,0\" /f >nul",
      "reg add \"HKCU\\Software\\Classes\\maio-open\\shell\\open\\command\" /ve /d \"wscript.exe \\\"%VBS_FILE%\\\" \\\"%%1\\\"\" /f >nul",
      "",
      "echo ========================================================",
      "echo  MAIO-Open installed successfully!",
      "echo  You can now close this window and return to MAIO web.",
      "echo ========================================================",
      "pause"
    ];
    return bat.join("\r\n");
  }

  downloadInstaller() {
    const content = this.buildInstallerBatContent();
    const blob = new Blob([content], { type: 'application/x-msdos-program' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'install-maio-open.bat';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  triggerUri(uri) {
    const a = document.createElement('a');
    a.href = uri;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => a.remove(), 1000);
  }

  openViaProtocol(action, localPath) {
    const clean = String(localPath || '').trim();
    if (!clean) return false;
    const safeAction = ['open', 'folder', 'select'].includes(action) ? action : 'open';
    const uri = `maio-open://${safeAction}?path=${encodeURIComponent(clean)}`;
    this.triggerUri(uri);
    return true;
  }

  async promptInstallModal({ actionLabel = 'Mở trực tiếp trên máy tính', onUseBrowser = null, onOpenProtocol = null } = {}) {
    return new Promise((resolve) => {
      const existing = document.getElementById('maio-open-install-modal');
      if (existing) existing.remove();

      const modal = document.createElement('div');
      modal.id = 'maio-open-install-modal';
      modal.className = 'modal-overlay';
      modal.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,0.55);backdrop-filter:blur(3px);z-index:10001;display:flex;align-items:center;justify-content:center;padding:16px;';

      modal.innerHTML = `
        <div style="background:var(--bg-card);color:var(--text-main);border:1px solid var(--border);border-radius:12px;width:100%;max-width:520px;box-shadow:0 20px 50px rgba(0,0,0,0.25);overflow:hidden;">
          <div style="padding:16px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;">
            <div style="font-weight:700;font-size:1.02rem;display:flex;align-items:center;gap:8px;">
              <i class="fa-solid fa-bolt" style="color:var(--primary);"></i>
              <span>Cài đặt MAIO-Open để mở file trực tiếp</span>
            </div>
            <button type="button" id="btn-maio-open-modal-close" class="button button-outline button-sm"><i class="fa-solid fa-xmark"></i></button>
          </div>
          <div style="padding:16px 20px;font-size:0.9rem;line-height:1.55;">
            <p style="margin:0 0 10px 0;">Để bấm <strong>${actionLabel}</strong> là bật thẳng Excel / Word / File Explorer trên máy Windows (không cần chuột phải → Open), bạn chỉ cần cài <strong>MAIO-Open</strong> 1 lần duy nhất:</p>
            <ol style="margin:0 0 14px 18px;padding:0;color:var(--text-muted);font-size:0.86rem;">
              <li>Bấm <strong>Tải bộ cài MAIO-Open (.bat)</strong> bên dưới và chạy file vừa tải (không cần quyền Admin).</li>
              <li>Sau khi cửa sổ cài đặt báo thành công, bấm <strong>Tôi đã cài xong</strong> (tick <em>Always allow</em> nếu trình duyệt hỏi).</li>
            </ol>
            <div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;">
              <button type="button" id="btn-maio-open-modal-download" class="button button-primary" style="flex:1;min-width:190px;">
                <i class="fa-solid fa-download"></i> Tải bộ cài MAIO-Open (.bat)
              </button>
              <button type="button" id="btn-maio-open-modal-installed" class="button button-outline" style="flex:1;min-width:170px;border-color:var(--success);color:var(--success);">
                <i class="fa-solid fa-check"></i> Tôi đã cài xong
              </button>
            </div>
          </div>
          <div style="padding:12px 20px;border-top:1px solid var(--border);background:var(--bg-app);display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;">
            <span style="font-size:0.78rem;color:var(--text-muted);">Bạn có thể bật/tắt lại trong mục Cài đặt bất cứ lúc nào.</span>
            <button type="button" id="btn-maio-open-modal-browser" class="button button-outline button-sm">
              Dùng chế độ trình duyệt (Không nhắc lại)
            </button>
          </div>
        </div>
      `;

      document.body.appendChild(modal);

      const close = (result) => {
        modal.remove();
        resolve(result);
      };

      modal.addEventListener('click', (e) => {
        if (e.target === modal) close('cancel');
      });
      modal.querySelector('#btn-maio-open-modal-close')?.addEventListener('click', () => close('cancel'));

      modal.querySelector('#btn-maio-open-modal-download')?.addEventListener('click', () => {
        this.downloadInstaller();
      });

      modal.querySelector('#btn-maio-open-modal-installed')?.addEventListener('click', () => {
        this.setStatus('installed');
        close('installed');
        if (typeof onOpenProtocol === 'function') onOpenProtocol();
      });

      modal.querySelector('#btn-maio-open-modal-browser')?.addEventListener('click', () => {
        this.setStatus('browser_only');
        close('browser_only');
        if (typeof onUseBrowser === 'function') onUseBrowser();
      });
    });
  }
}

export const maioOpen = new MaioOpenHelper();
