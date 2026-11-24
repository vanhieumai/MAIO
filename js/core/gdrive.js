/**
 * MAIO Google Drive Service
 * Quản lý xác thực OAuth 2.0 (Google Identity Services) và đồng bộ tệp tin qua Google Drive API v3.
 */

// Cấu hình dự phòng mặc định nếu file gdrive-config.js chưa kịp tải hoặc bị chặn cache
const DEFAULT_GDRIVE_CONFIG = {
  clientId: "429965167454-7p2a5fodhue357s10chcussc8u45ufn0.apps.googleusercontent.com",
  apiKey: "AIzaSyAVqM4ADNcwXo-43eV-qf-oNaCHpD26qqo",
  appFolderName: "MAIO_Data",
  scopes: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email"
};

export class GoogleDriveService {
  constructor(config = window.MAIO_GDRIVE_CONFIG) {
    this.config = Object.assign({}, DEFAULT_GDRIVE_CONFIG, config || window.MAIO_GDRIVE_CONFIG || {});
    this.tokenClient = null;
    this.accessToken = null;
    this.currentUser = null;
    this.appFolderId = null;
    this.isInitialized = false;
    this.listeners = [];
  }

  /**
   * Đăng ký lắng nghe thay đổi trạng thái đăng nhập
   */
  onAuthStateChanged(callback) {
    this.listeners.push(callback);
    if (this.currentUser) {
      callback(this.currentUser);
    }
  }

  notifyAuthChange(user) {
    this.currentUser = user;
    this.listeners.forEach((cb) => cb(user));
  }

  /**
   * Khởi tạo Google Identity Services
   */
  async init() {
    if (this.isInitialized) return true;

    if (window.MAIO_GDRIVE_CONFIG) {
      this.config = Object.assign({}, DEFAULT_GDRIVE_CONFIG, this.config, window.MAIO_GDRIVE_CONFIG);
    }

    return new Promise((resolve, reject) => {
      const checkScript = () => {
        if (window.google && window.google.accounts && window.google.accounts.oauth2) {
          try {
            this.tokenClient = window.google.accounts.oauth2.initTokenClient({
              client_id: this.config.clientId,
              scope: this.config.scopes,
              callback: async (tokenResponse) => {
                if (tokenResponse && tokenResponse.access_token) {
                  // Kiểm tra xem đã được tích chọn cấp quyền Drive hay chưa
                  const hasDriveScope = window.google.accounts.oauth2.hasGrantedAnyScope(
                    tokenResponse,
                    'https://www.googleapis.com/auth/drive.file'
                  );

                  if (!hasDriveScope) {
                    alert('Lưu ý: Bạn chưa tích chọn ô cho phép MAIO tạo và quản lý tệp trên Google Drive. Vui lòng đăng nhập lại và tích chọn ô này để có thể đồng bộ dữ liệu.');
                    this.signOut();
                    return;
                  }

                  this.accessToken = tokenResponse.access_token;
                  sessionStorage.setItem('maio_gdrive_token', this.accessToken);
                  await this.fetchUserInfo();
                  if (this.currentUser) {
                    localStorage.setItem('maio_gdrive_user', JSON.stringify(this.currentUser));
                  }
                  try {
                    await this.ensureAppFolder();
                  } catch (err) {
                    console.warn('Chưa tạo được thư mục:', err);
                  }
                  this.notifyAuthChange(this.currentUser);
                }
              },
            });

            // Khôi phục phiên người dùng nếu đã đăng nhập trước đó
            const savedUserJson = localStorage.getItem('maio_gdrive_user');
            const savedToken = sessionStorage.getItem('maio_gdrive_token');
            if (savedUserJson) {
              try {
                const savedUser = JSON.parse(savedUserJson);
                if (savedUser && savedUser.email) {
                  this.currentUser = savedUser;
                  this.accessToken = savedToken || 'saved_session';
                  this.notifyAuthChange(this.currentUser);
                  if (!savedUser.isOffline && savedToken) {
                    this.ensureAppFolder().catch(() => {});
                  }
                } else {
                  this.notifyAuthChange(null);
                }
              } catch (e) {
                localStorage.removeItem('maio_gdrive_user');
                this.notifyAuthChange(null);
              }
            } else if (savedToken) {
              this.accessToken = savedToken;
              this.fetchUserInfo()
                .then(() => {
                  if (this.currentUser) {
                    localStorage.setItem('maio_gdrive_user', JSON.stringify(this.currentUser));
                  }
                  return this.ensureAppFolder();
                })
                .then(() => this.notifyAuthChange(this.currentUser))
                .catch(() => {
                  this.signOut();
                });
            } else {
              // Chưa từng đăng nhập -> kích hoạt thông báo Auth Gate ngay
              this.notifyAuthChange(null);
            }

            this.isInitialized = true;
            resolve(true);
          } catch (err) {
            this.initError = err;
            console.error('[MAIO Google Drive] initTokenClient Error:', err);
            this.notifyAuthChange(null);
            reject(err);
          }
        } else {
          scriptAttempts++;
          if (scriptAttempts > 60) {
            this.initError = new Error('Không thể tải thư viện Google Identity Services từ accounts.google.com sau 6 giây (có thể bị trình duyệt chặn hoặc mất mạng).');
            this.notifyAuthChange(null);
            resolve(false);
          } else {
            setTimeout(checkScript, 100);
          }
        }
      };

      let scriptAttempts = 0;
      checkScript();
    });
  }

  /**
   * Mở popup đăng nhập tài khoản Google
   */
  async signIn() {
    if (!this.tokenClient) {
      // Thử khởi tạo lại nếu script Google Identity Services tải chậm
      try {
        await this.init();
      } catch (e) {
        this.initError = e;
        console.warn('[MAIO Google Drive] Lỗi khởi tạo GIS:', e);
      }
    }

    if (!this.tokenClient) {
      const currentOrigin = window.location.origin;
      const detailError = this.initError ? `\n\nChi tiết lỗi kỹ thuật: ${this.initError.message || this.initError}` : '';

      throw new Error(
        `Google Identity Services chưa được khởi tạo thành công!${detailError}\n\n` +
        `Hướng dẫn khắc phục:\n` +
        `1. Bạn đã thêm email vào "Test users", nhưng BẮT BUỘC phải thêm URL website vào mục "Authorized JavaScript origins":\n` +
        `   • Truy cập: console.cloud.google.com -> APIs & Services -> Credentials\n` +
        `   • Bấm vào OAuth 2.0 Client ID của bạn\n` +
        `   • Cuộn xuống mục "Authorized JavaScript origins", bấm "+ ADD URI" và điền:\n` +
        `     ${currentOrigin}\n` +
        `   • Bấm nút SAVE (Lưu) ở cuối trang và đợi 1 - 2 phút.\n\n` +
        `2. Nếu dùng trình duyệt chặn quảng cáo (Brave, AdBlock), vui lòng tạm tắt để cho phép tải accounts.google.com.\n\n` +
        `(Mẹo: Bạn có thể bấm "Làm việc Cục bộ" để chọn file CPMS/Excel dùng ngay mà không cần kết nối Google Drive).`
      );
    }
    this.tokenClient.requestAccessToken({ prompt: 'consent' });
  }

  /**
   * Đăng nhập chế độ Cục bộ (Offline) - lưu trữ trực tiếp trên máy tính
   */
  signInOffline(email = 'local_user@maio.local', name = 'Người dùng Cục bộ (Offline)') {
    this.currentUser = {
      email: 'local_user@maio.local',
      name: 'Người dùng Cục bộ (Offline)',
      picture: '',
      isOffline: true,
    };
    this.accessToken = 'offline_local_session';
    localStorage.setItem('maio_gdrive_user', JSON.stringify(this.currentUser));
    sessionStorage.setItem('maio_gdrive_token', this.accessToken);
    this.notifyAuthChange(this.currentUser);
    return this.currentUser;
  }

  /**
   * Đăng xuất và xóa token phiên
   */
  signOut() {
    if (this.accessToken && window.google?.accounts?.oauth2 && !this.currentUser?.isOffline) {
      try {
        window.google.accounts.oauth2.revoke(this.accessToken, () => {});
      } catch (e) {
        console.warn('Lỗi thu hồi token Google:', e);
      }
    }
    this.accessToken = null;
    this.currentUser = null;
    this.appFolderId = null;
    sessionStorage.removeItem('maio_gdrive_token');
    localStorage.removeItem('maio_gdrive_user');
    this.notifyAuthChange(null);
  }

  isSignedIn() {
    return !!this.accessToken;
  }

  /**
   * Lấy thông tin tài khoản người dùng từ Google
   */
  async fetchUserInfo() {
    if (!this.accessToken) return null;
    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${this.accessToken}` },
      });
      if (res.ok) {
        this.currentUser = await res.json();
        return this.currentUser;
      }
    } catch (e) {
      console.warn('Không thể lấy thông tin người dùng Google:', e);
    }
    return null;
  }

  /**
   * Tìm hoặc tự động tạo thư mục MAIO_Data trên Google Drive
   */
  async ensureAppFolder() {
    if (this.appFolderId) return this.appFolderId;
    if (!this.accessToken) return null;

    const folderName = this.config.appFolderName || 'MAIO_Data';
    const query = `mimeType='application/vnd.google-apps.folder' and name='${folderName}' and trashed=false`;

    try {
      const searchRes = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)`,
        {
          headers: { Authorization: `Bearer ${this.accessToken}` },
        }
      );

      const searchData = await searchRes.json();
      if (searchData.files && searchData.files.length > 0) {
        this.appFolderId = searchData.files[0].id;
        return this.appFolderId;
      }

      // Chưa có thì tạo thư mục mới
      const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: folderName,
          mimeType: 'application/vnd.google-apps.folder',
        }),
      });

      if (!createRes.ok) {
        let errMsg = '';
        try {
          const errData = await createRes.json();
          errMsg = errData?.error?.message || JSON.stringify(errData);
        } catch (_) {
          errMsg = await createRes.text();
        }
        throw new Error(`Không thể tạo thư mục [${createRes.status}]: ${errMsg}`);
      }

      const createData = await createRes.json();
      this.appFolderId = createData.id;
      return this.appFolderId;
    } catch (err) {
      console.error('Lỗi khi kiểm tra thư mục MAIO trên Google Drive:', err);
      throw err;
    }
  }

  /**
   * Tìm một thư mục con theo tên trong thư mục cha
   */
  async findFolder(folderName, parentFolderId) {
    if (!this.accessToken) return null;
    const parentId = parentFolderId || this.appFolderId;
    const query = `mimeType='application/vnd.google-apps.folder' and name='${folderName}' and '${parentId}' in parents and trashed=false`;
    try {
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)`,
        { headers: { Authorization: `Bearer ${this.accessToken}` } }
      );
      if (!res.ok) return null;
      const data = await res.json();
      return data.files && data.files.length > 0 ? data.files[0].id : null;
    } catch (_) {
      return null;
    }
  }

  /**
   * Tạo một thư mục con trong thư mục cha
   */
  async createFolder(folderName, parentFolderId) {
    if (!this.accessToken) return null;
    const parentId = parentFolderId || this.appFolderId;
    try {
      const res = await fetch('https://www.googleapis.com/drive/v3/files', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: folderName,
          mimeType: 'application/vnd.google-apps.folder',
          parents: [parentId],
        }),
      });
      if (!res.ok) {
        throw new Error(`Không thể tạo thư mục "${folderName}" [${res.status}]`);
      }
      const data = await res.json();
      return data.id;
    } catch (e) {
      console.warn('Lỗi tạo thư mục Drive:', e);
      throw e;
    }
  }

  /**
   * Đảm bảo một chuỗi thư mục lồng nhau tồn tại trên Google Drive
   * Vd: ['01_QuanLy_DuAn', 'DuAn_CHC1', 'CPMS_Data']
   */
  async ensureFolderPath(pathSegments, rootFolderId = null) {
    await this.ensureAppFolder();
    let currentParent = rootFolderId || this.appFolderId;
    if (!currentParent) return null;

    for (const segment of pathSegments) {
      if (!segment) continue;
      let folderId = await this.findFolder(segment, currentParent);
      if (!folderId) {
        folderId = await this.createFolder(segment, currentParent);
      }
      currentParent = folderId;
    }
    return currentParent;
  }

  /**
   * Danh sách các file dự án lưu trong thư mục MAIO_Data hoặc thư mục chỉ định
   */
  async listProjectFiles(targetFolderId = null) {
    await this.ensureAppFolder();
    const parentId = targetFolderId || this.appFolderId;
    if (!parentId) return [];

    const query = `'${parentId}' in parents and trashed=false`;
    const res = await fetch(
      `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(
        query
      )}&fields=files(id,name,mimeType,modifiedTime,size)&orderBy=modifiedTime desc`,
      {
        headers: { Authorization: `Bearer ${this.accessToken}` },
      }
    );

    if (!res.ok) {
      let errDetail = '';
      try {
        const errJson = await res.json();
        errDetail = errJson?.error?.message || JSON.stringify(errJson);
      } catch (_) {
        errDetail = await res.text();
      }
      throw new Error(`Lỗi tải danh sách file [${res.status}]: ${errDetail}`);
    }

    const data = await res.json();
    return data.files || [];
  }

  /**
   * Đọc nội dung file JSON từ Google Drive
   */
  async readJsonFile(fileId) {
    const res = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });
    if (!res.ok) {
      let errDetail = '';
      try {
        const errJson = await res.json();
        errDetail = errJson?.error?.message || JSON.stringify(errJson);
      } catch (_) {
        errDetail = await res.text();
      }
      throw new Error(`Lỗi đọc file [${res.status}]: ${errDetail}`);
    }
    return await res.json();
  }

  /**
   * Lưu hoặc ghi đè file JSON lên thư mục MAIO_Data hoặc targetFolderId
   */
  async saveJsonFile(fileName, dataObj, targetFolderId = null) {
    await this.ensureAppFolder();
    const parentId = targetFolderId || this.appFolderId;
    const files = await this.listProjectFiles(parentId);
    const existingFile = files.find((f) => f.name === fileName);

    const metadata = {
      name: fileName,
      mimeType: 'application/json',
      ...(!existingFile && { parents: [parentId] }),
    };

    const fileContent = JSON.stringify(dataObj, null, 2);
    const boundary = '-------314159265358979323846';
    const delimiter = `\r\n--${boundary}\r\n`;
    const closeDelimiter = `\r\n--${boundary}--`;

    const bodyParts = [
      delimiter,
      'Content-Type: application/json; charset=UTF-8\r\n\r\n',
      JSON.stringify(metadata),
      delimiter,
      'Content-Type: application/json; charset=UTF-8\r\n\r\n',
      fileContent,
      closeDelimiter,
    ];

    const multipartBlob = new Blob(bodyParts, { type: `multipart/related; boundary=${boundary}` });

    let url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
    let method = 'POST';

    if (existingFile) {
      url = `https://www.googleapis.com/upload/drive/v3/files/${existingFile.id}?uploadType=multipart`;
      method = 'PATCH';
    }

    const res = await fetch(url, {
      method: method,
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
      },
      body: multipartBlob,
    });

    if (!res.ok) {
      let errDetail = '';
      try {
        const errJson = await res.json();
        errDetail = errJson?.error?.message || JSON.stringify(errJson);
      } catch (_) {
        errDetail = await res.text();
      }
      throw new Error(`Lỗi lưu file lên Google Drive [${res.status}]: ${errDetail || res.statusText}`);
    }
    return await res.json();
  }

  /**
   * Tải nội dung file nhị phân (Blob) từ Google Drive
   */
  async downloadBlobFile(fileId) {
    if (!this.accessToken) throw new Error('Chưa đăng nhập Google Drive.');
    const res = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });
    if (!res.ok) throw new Error(`Lỗi tải file nhị phân [${res.status}]`);
    return await res.blob();
  }

  /**
   * Liệt kê tất cả file và folder con trong một thư mục Drive
   */
  async listFolderContents(folderId) {
    if (!this.accessToken) return [];
    const query = `'${folderId}' in parents and trashed=false`;
    try {
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,mimeType,size,modifiedTime)&pageSize=1000`,
        { headers: { Authorization: `Bearer ${this.accessToken}` } }
      );
      if (!res.ok) return [];
      const data = await res.json();
      return data.files || [];
    } catch (_) {
      return [];
    }
  }

  /**
   * Tìm tất cả file snapshot dự án trên Google Drive
   */
  async listProjectSnapshots() {
    if (!this.accessToken) return [];
    await this.ensureAppFolder();
    const query = `name contains 'MAIO_CPMS_Snapshot_' and mimeType='application/json' and trashed=false`;
    try {
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,modifiedTime,parents)&orderBy=modifiedTime desc`,
        { headers: { Authorization: `Bearer ${this.accessToken}` } }
      );
      if (!res.ok) return [];
      const data = await res.json();
      return data.files || [];
    } catch (e) {
      console.warn('Lỗi tìm snapshot dự án:', e);
      return [];
    }
  }

  /**
   * Tải đệ quy toàn bộ thư mục từ Drive về thư mục máy tính (FileSystemDirectoryHandle)
   */
  async downloadDriveFolderToLocal(driveFolderId, localDirHandle) {
    const items = await this.listFolderContents(driveFolderId);
    let downloadedCount = 0;

    for (const item of items) {
      if (item.mimeType === 'application/vnd.google-apps.folder') {
        const subDirHandle = await localDirHandle.getDirectoryHandle(item.name, { create: true });
        downloadedCount += await this.downloadDriveFolderToLocal(item.id, subDirHandle);
      } else {
        try {
          const blob = await this.downloadBlobFile(item.id);
          const fileHandle = await localDirHandle.getFileHandle(item.name, { create: true });
          const writable = await fileHandle.createWritable();
          await writable.write(blob);
          await writable.close();
          downloadedCount++;
        } catch (err) {
          console.warn(`Lỗi ghi file "${item.name}" về local:`, err);
        }
      }
    }
    return downloadedCount;
  }
}

if (!window.__MAIO_GDRIVE_SERVICE__) {
  window.__MAIO_GDRIVE_SERVICE__ = new GoogleDriveService();
}
export const gdrive = window.__MAIO_GDRIVE_SERVICE__;
