/**
 * 煤炭评标系统配置文件
 * 部署前请填写正确的值
 */

const CONFIG = {
  // Google Apps Script Web App URL
  SHEETS_API_URL: 'https://script.google.com/a/macros/kchenergy.com/s/AKfycbzsoTX69TvbAOUzJeHtICJkC20gGuv6r2E6-wjtyQZKPMjUtTxDILUdd_j2E454Xevu/exec',
  
  // Google Sheet ID
  SHEET_ID: '1i9Mhos5GLV21MrZyu0VCM-HKvAvz8O-TzG5VtwmzGnk',
  
  // 是否启用云端同步
  CLOUD_ENABLED: true,
  
  // 本地存储键名前缀
  LOCAL_STORAGE_PREFIX: 'kch_coal_',
  
  // 自动保存间隔（毫秒）
  AUTO_SAVE_INTERVAL: 30000,
  
  // 调试模式
  DEBUG: false
};

// 检查配置是否完整
function checkConfig() {
  if (CONFIG.CLOUD_ENABLED) {
    if (!CONFIG.SHEETS_API_URL) {
      console.warn('⚠️ SHEETS_API_URL 未配置，云端功能将不可用');
      return false;
    }
    if (!CONFIG.SHEET_ID) {
      console.warn('⚠️ SHEET_ID 未配置，云端功能将不可用');
      return false;
    }
  }
  return true;
}
