/**
 * 煤炭评标系统配置文件
 * 部署前请填写正确的值
 */

const CONFIG = {
  // Google Apps Script Web App URL (已更新为你的最新链接)
  SHEETS_API_URL: 'https://script.google.com/macros/s/AKfycbwmQ-zZtEyV19OXb_3Y2c_CIQBVmKfwkgQRzoQjDbWkSDudqcTbcTsopS2HBRzPwyRx/exec',
  
  // Google Sheet ID
  SHEET_ID: '1i9Mhos5GLV21MrZyu0VCM-HKvAvz8O-TzG5VtwmzGnk',
  
  // 是否启用云端同步
  CLOUD_ENABLED: true,
  
  // 本地存储键名前缀
  LOCAL_STORAGE_PREFIX: 'kch_coal_',
  
  // 自动保存间隔（毫秒）
  AUTO_SAVE_INTERVAL: 30000,
  
  // 调试模式 (我帮你把这里改成了 true，这样如果有问题，按 F12 可以在控制台看到更详细的报错)
  DEBUG: true
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