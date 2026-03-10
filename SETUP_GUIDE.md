# 煤炭评标系统 - Google Sheets 部署指南

## 架构说明

```
┌─────────────────────────────────────────────────────────┐
│                    Google Sheets                         │
├─────────────────────────────────────────────────────────┤
│  📊 ICI_Baseline        - ICI基准价格历史               │
│  📊 MineDatabase        - 矿山数据库（共享）             │
│  📊 OfficialBids        - 正式评标记录 ⚠️ 仅管理员      │
│  📊 PersonalDrafts      - 个人草稿（按用户名分区）       │
│  📊 AuditLog            - 操作日志                       │
└─────────────────────────────────────────────────────────┘
```

## 步骤 1: 创建 Google Sheet

1. 打开 Google Sheets: https://sheets.google.com
2. 创建新表格，命名为: `KCHenergy 煤炭评标系统`
3. 复制表格ID（URL中 `/d/` 和 `/edit` 之间的部分）
   例如: `https://docs.google.com/spreadsheets/d/1ABC...XYZ/edit`
   表格ID就是: `1ABC...XYZ`

## 步骤 2: 创建工作表

创建以下5个工作表（Sheet tabs）:

### Sheet 1: `ICI_Baseline`
| 列 | 内容 |
|----|------|
| A | 日期 |
| B | ICI5000价格 |
| C | NAR |
| D | 硫分 |
| E | 灰分 |
| F | 氮含量 |
| G | 基准运费 |
| H | 税费 |
| I | 更新人 |

### Sheet 2: `MineDatabase`
| 列 | 内容 |
|----|------|
| A | 矿山ID |
| B | 矿山名称 |
| C | 供应商 |
| D | 典型NAR |
| E | 典型硫分 |
| F | 典型灰分 |
| G | 典型氮含量 |
| H | 备注 |

### Sheet 3: `OfficialBids`
| 列 | 内容 |
|----|------|
| A | 批次ID |
| B | 批次名称 |
| C | 创建日期 |
| D | Genco |
| E | ICI基准评标值 |
| F | 最优矿山 |
| G | 最优评标值 |
| H | 完整数据(JSON) |
| I | 创建人 |
| J | 状态(draft/final) |

### Sheet 4: `PersonalDrafts`
| 列 | 内容 |
|----|------|
| A | 草稿ID |
| B | 用户邮箱 |
| C | 草稿名称 |
| D | 创建日期 |
| E | 完整数据(JSON) |

### Sheet 5: `AuditLog`
| 列 | 内容 |
|----|------|
| A | 时间戳 |
| B | 用户邮箱 |
| C | 操作类型 |
| D | 详情 |

## 步骤 3: 部署 Apps Script API

1. 在 Google Sheet 中，点击 **扩展程序 → Apps Script**
2. 删除默认代码，粘贴 `google-apps-script.js` 的内容
3. 点击 **部署 → 新建部署**
4. 选择类型: **网页应用**
5. 设置:
   - 描述: `Coal Bidding API v1`
   - 执行身份: `我`
   - 访问权限: `任何人`（或 `仅限组织内`）
6. 点击 **部署**，复制生成的 Web App URL

## 步骤 4: 配置权限

### 管理员设置
在 Apps Script 代码顶部修改 `ADMIN_EMAILS`:
```javascript
const ADMIN_EMAILS = [
  'lechen@kchenergy.com',    // 你的邮箱
  'another.admin@kchenergy.com'  // 另一位管理员
];
```

### Sheet 共享设置
1. 点击 Sheet 右上角 **共享**
2. 添加团队10人的邮箱
3. 权限设置为 **编辑者**（API会在代码层面控制写入权限）

## 步骤 5: 配置前端

打开 `config.js`，填入:
```javascript
const CONFIG = {
  SHEETS_API_URL: '你的 Web App URL',
  SHEET_ID: '你的表格ID'
};
```

## 步骤 6: 部署前端

### 选项 A: Google Sites（推荐）
1. 打开 https://sites.google.com
2. 创建新网站
3. 插入 → 嵌入 → 嵌入代码
4. 上传 HTML/JS/CSS 文件

### 选项 B: GitHub Pages（免费）
1. 创建 GitHub 仓库
2. 上传所有文件
3. 启用 GitHub Pages

### 选项 C: 公司内网服务器
将文件放到任何 HTTP 服务器即可

---

## 权限矩阵

| 功能 | 普通用户 | 管理员 |
|------|----------|--------|
| 查看ICI基准 | ✅ | ✅ |
| 更新ICI基准 | ❌ | ✅ |
| 查看矿山数据库 | ✅ | ✅ |
| 添加/修改矿山 | ✅ | ✅ |
| 创建个人草稿 | ✅ | ✅ |
| 查看个人草稿 | ✅ 仅自己 | ✅ 全部 |
| 创建正式评标 | ❌ | ✅ |
| 修改正式评标 | ❌ | ✅ |
| 查看正式评标 | ✅ | ✅ |
| 导出Excel | ✅ | ✅ |
