/**
 * 煤炭评标系统 - API 通信层
 * 处理与 Google Sheets 的数据交互
 */

class CoalBiddingAPI {
  constructor() {
    this.baseUrl = CONFIG.SHEETS_API_URL;
    this.enabled = CONFIG.CLOUD_ENABLED && checkConfig();
    this.userEmail = null;
    this.userInfo = null;
  }

  // ==================== 基础请求 ====================

  async request(action, data = {}) {
    if (!this.enabled) {
      console.log('云端未启用，使用本地存储');
      return { success: false, error: 'Cloud not enabled' };
    }

    try {
      const params = new URLSearchParams({
        action: action,
        userEmail: this.userEmail || '',
        ...data
      });

      if (CONFIG.DEBUG) {
        console.log('API Request:', action, data);
      }

      const response = await fetch(`${this.baseUrl}?${params.toString()}`, {
        method: 'GET',
        mode: 'cors'
      });

      const result = await response.json();

      if (CONFIG.DEBUG) {
        console.log('API Response:', result);
      }

      return result;
    } catch (error) {
      console.error('API Error:', error);
      return { success: false, error: error.message };
    }
  }

  async postRequest(action, data) {
    if (!this.enabled) {
      return { success: false, error: 'Cloud not enabled' };
    }

    try {
      const params = new URLSearchParams({
        action: action,
        userEmail: this.userEmail || '',
        data: JSON.stringify(data)
      });

      const response = await fetch(`${this.baseUrl}?${params.toString()}`, {
        method: 'POST',
        mode: 'cors'
      });

      return await response.json();
    } catch (error) {
      console.error('API Error:', error);
      return { success: false, error: error.message };
    }
  }

  // ==================== 用户认证 ====================

  async initialize(email) {
    this.userEmail = email;
    
    if (!this.enabled) {
      // 本地模式，模拟管理员
      this.userInfo = {
        email: email,
        isAdmin: true,
        permissions: {
          canEditICI: true,
          canCreateOfficialBid: true,
          canEditOfficialBid: true,
          canEditMineDB: true,
          canCreateDraft: true
        }
      };
      return { success: true, data: this.userInfo };
    }

    const result = await this.request('getUserInfo');
    if (result.success) {
      this.userInfo = result.data;
    }
    return result;
  }

  isAdmin() {
    return this.userInfo?.isAdmin || false;
  }

  canEditICI() {
    return this.userInfo?.permissions?.canEditICI || false;
  }

  canCreateOfficialBid() {
    return this.userInfo?.permissions?.canCreateOfficialBid || false;
  }

  // ==================== ICI 基准 ====================

  async getICIBaseline() {
    if (!this.enabled) {
      return this.getLocalICIBaseline();
    }
    return await this.request('getICIBaseline');
  }

  async updateICIBaseline(data) {
    if (!this.enabled) {
      return this.saveLocalICIBaseline(data);
    }
    return await this.postRequest('updateICIBaseline', data);
  }

  getLocalICIBaseline() {
    const saved = localStorage.getItem(CONFIG.LOCAL_STORAGE_PREFIX + 'ici_baseline');
    return {
      success: true,
      data: saved ? JSON.parse(saved) : null
    };
  }

  saveLocalICIBaseline(data) {
    data.date = new Date().toISOString();
    data.updatedBy = this.userEmail || 'local';
    localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'ici_baseline', JSON.stringify(data));
    return { success: true };
  }

  // ==================== 矿山数据库 ====================

  async getMineDatabase() {
    if (!this.enabled) {
      return this.getLocalMineDatabase();
    }
    return await this.request('getMineDatabase');
  }

  async addMine(data) {
    if (!this.enabled) {
      return this.addLocalMine(data);
    }
    return await this.postRequest('addMine', data);
  }

  async updateMine(data) {
    if (!this.enabled) {
      return this.updateLocalMine(data);
    }
    return await this.postRequest('updateMine', data);
  }

  getLocalMineDatabase() {
    const saved = localStorage.getItem(CONFIG.LOCAL_STORAGE_PREFIX + 'mine_database');
    return {
      success: true,
      data: saved ? JSON.parse(saved) : []
    };
  }

  addLocalMine(data) {
    const result = this.getLocalMineDatabase();
    const mines = result.data || [];
    data.id = Date.now().toString();
    mines.push(data);
    localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'mine_database', JSON.stringify(mines));
    return { success: true, data: { id: data.id } };
  }

  updateLocalMine(data) {
    const result = this.getLocalMineDatabase();
    const mines = result.data || [];
    const index = mines.findIndex(m => m.id === data.id);
    if (index >= 0) {
      mines[index] = { ...mines[index], ...data };
      localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'mine_database', JSON.stringify(mines));
      return { success: true };
    }
    return { success: false, error: '矿山未找到' };
  }

  // ==================== 正式评标 ====================

  async getOfficialBids() {
    if (!this.enabled) {
      return this.getLocalOfficialBids();
    }
    return await this.request('getOfficialBids');
  }

  async createOfficialBid(data) {
    if (!this.enabled) {
      return this.createLocalOfficialBid(data);
    }
    return await this.postRequest('createOfficialBid', data);
  }

  async updateOfficialBid(data) {
    if (!this.enabled) {
      return this.updateLocalOfficialBid(data);
    }
    return await this.postRequest('updateOfficialBid', data);
  }

  getLocalOfficialBids() {
    const saved = localStorage.getItem(CONFIG.LOCAL_STORAGE_PREFIX + 'official_bids');
    return {
      success: true,
      data: saved ? JSON.parse(saved) : []
    };
  }

  createLocalOfficialBid(data) {
    const result = this.getLocalOfficialBids();
    const bids = result.data || [];
    data.id = Date.now().toString();
    data.createdAt = new Date().toISOString();
    data.createdBy = this.userEmail || 'local';
    bids.push(data);
    localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'official_bids', JSON.stringify(bids));
    return { success: true, data: { id: data.id } };
  }

  updateLocalOfficialBid(data) {
    const result = this.getLocalOfficialBids();
    const bids = result.data || [];
    const index = bids.findIndex(b => b.id === data.id);
    if (index >= 0) {
      bids[index] = { ...bids[index], ...data };
      localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'official_bids', JSON.stringify(bids));
      return { success: true };
    }
    return { success: false, error: '评标记录未找到' };
  }

  // ==================== 个人草稿 ====================

  async getMyDrafts() {
    if (!this.enabled) {
      return this.getLocalDrafts();
    }
    return await this.request('getMyDrafts');
  }

  async saveDraft(data) {
    if (!this.enabled) {
      return this.saveLocalDraft(data);
    }
    return await this.postRequest('saveDraft', data);
  }

  async deleteDraft(draftId) {
    if (!this.enabled) {
      return this.deleteLocalDraft(draftId);
    }
    return await this.request('deleteDraft', { draftId });
  }

  getLocalDrafts() {
    const saved = localStorage.getItem(CONFIG.LOCAL_STORAGE_PREFIX + 'drafts');
    return {
      success: true,
      data: saved ? JSON.parse(saved) : []
    };
  }

  saveLocalDraft(data) {
    const result = this.getLocalDrafts();
    const drafts = result.data || [];
    
    if (data.id) {
      const index = drafts.findIndex(d => d.id === data.id);
      if (index >= 0) {
        drafts[index] = { ...drafts[index], ...data, updatedAt: new Date().toISOString() };
        localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'drafts', JSON.stringify(drafts));
        return { success: true, data: { id: data.id } };
      }
    }
    
    data.id = Date.now().toString();
    data.createdAt = new Date().toISOString();
    data.updatedAt = data.createdAt;
    data.userEmail = this.userEmail || 'local';
    drafts.push(data);
    localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'drafts', JSON.stringify(drafts));
    return { success: true, data: { id: data.id } };
  }

  deleteLocalDraft(draftId) {
    const result = this.getLocalDrafts();
    const drafts = result.data || [];
    const filtered = drafts.filter(d => d.id !== draftId);
    localStorage.setItem(CONFIG.LOCAL_STORAGE_PREFIX + 'drafts', JSON.stringify(filtered));
    return { success: true };
  }

  // ==================== 同步状态 ====================

  isCloudEnabled() {
    return this.enabled;
  }

  getConnectionStatus() {
    return {
      cloudEnabled: this.enabled,
      userEmail: this.userEmail,
      isAdmin: this.isAdmin()
    };
  }
}

// 全局 API 实例
const api = new CoalBiddingAPI();
