/**
 * MAIO Google Drive Service
 * Quản lý xác thực OAuth 2.0 (Google Identity Services) và đồng bộ tệp tin qua Google Drive API v3.
 */

export class GoogleDriveService {
  constructor(config = window.MAIO_GDRIVE_CONFIG) {
    this.config = config || {};
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
                  try {
                    await this.ensureAppFolder();
                  } catch (err) {
                    console.warn('Chưa tạo được thư mục:', err);
                  }
                  this.notifyAuthChange(this.currentUser);
                }
              },
            });

            // Khôi phục token nếu còn hiệu lực trong phiên
            const savedToken = sessionStorage.getItem('maio_gdrive_token');
            if (savedToken) {
              this.accessToken = savedToken;
              this.fetchUserInfo()
                .then(() => this.ensureAppFolder())
                .then(() => this.notifyAuthChange(this.currentUser))
                .catch(() => {
                  this.signOut();
                });
            }

            this.isInitialized = true;
            resolve(true);
          } catch (err) {
            reject(err);
          }
        } else {
          setTimeout(checkScript, 100);
        }
      };

      checkScript();
    });
  }

  /**
   * Mở popup đăng nhập tài khoản Google
   */
  signIn() {
    if (!this.tokenClient) {
      throw new Error('Google Identity Services chưa được khởi tạo. Vui lòng kiểm tra Client ID.');
    }
    this.tokenClient.requestAccessToken({ prompt: 'consent' });
  }

  /**
   * Đăng xuất và xóa token phiên
   */
  signOut() {
    if (this.accessToken && window.google?.accounts?.oauth2) {
      window.google.accounts.oauth2.revoke(this.accessToken, () => {});
    }
    this.accessToken = null;
    this.currentUser = null;
    this.appFolderId = null;
    sessionStorage.removeItem('maio_gdrive_token');
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
   * Danh sách các file dự án lưu trong thư mục MAIO_Data
   */
  async listProjectFiles() {
    await this.ensureAppFolder();
    if (!this.appFolderId) return [];

    const query = `'${this.appFolderId}' in parents and trashed=false`;
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
   * Lưu hoặc ghi đè file JSON lên thư mục MAIO_Data
   */
  async saveJsonFile(fileName, dataObj) {
    await this.ensureAppFolder();
    const files = await this.listProjectFiles();
    const existingFile = files.find((f) => f.name === fileName);

    const metadata = {
      name: fileName,
      mimeType: 'application/json',
      ...(!existingFile && { parents: [this.appFolderId] }),
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
}

export const gdrive = new GoogleDriveService();
