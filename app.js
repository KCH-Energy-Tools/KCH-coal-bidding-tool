/**
 * 발전소 석탄 입찰 평가 시스템 - 메인 애플리케이션 로직
 * KCHenergy Internal Tool
 * Version 2.2 - 정확한 평가 공식 (Evaluation & 평가상)
 * 
 * 공식:
 * Evaluation = (CFR + Quality Penalty + P/C + Consumption Tax) × (6080 / NAR)
 * 평가상 = Bidder Evaluation - ICI-5000 기준 Evaluation
 */

// ==================== 사용자 권한 설정 ====================

//Authorized users - only these emails can log in
// admin: can create/edit official bids and modify ICI baseline
// user: can only view and create personal drafts
const AUTHORIZED_USERS = {
    'lechen@kchenergy.com': { name: 'Lechen', role: 'admin' },
    'eunn29@kchenergy.com': { name: 'Eun', role: 'user' },
    'jinwooook@kchenergy.com': { name: 'Jinwoo', role: 'user' },
    'mjk@kchenergy.com': { name: 'MJ', role: 'user' },
    'jseo@kchenergy.com': { name: 'JSeo', role: 'user' },
    'sw_jeong@kchenergy.com': { name: 'SW_Jeong', role: 'user' },
    'jsi@kchenergy.com': { name: 'JSi', role: 'user' },
    'ojw5710@kchenergy.com': { name: 'Ojw', role: 'user' },
    'thankbit@kchenergy.com': { name: 'Thankbit', role: 'user' },
    'khh@kchenergy.com': { name: 'KHH', role: 'user' },
};

function isAuthorized(email) {
    return AUTHORIZED_USERS[email.toLowerCase()] !== undefined;
}

function getUserRole(email) {
    const user = AUTHORIZED_USERS[email.toLowerCase()];
    return user ? user.role : null;
}

function getUserName(email) {
    const user = AUTHORIZED_USERS[email.toLowerCase()];
    return user ? user.name : email.split('@')[0];
}

// ==================== 발전사별 페널티 계수 ====================

const GENCO_COEFFICIENTS = {
    kospo:   { sulfur: 2.92, ash: 0.07,  nitrogen: 0.25,  pc: 0.61, carbon: 0,    name: 'KOSPO (남부발전)' },
    koen:    { sulfur: 4.53, ash: 0.09,  nitrogen: 0.39,  pc: 1.88, carbon: 0,    name: 'KOEN (남동발전)' },
    kowepo:  { sulfur: 3.53, ash: 0.043, nitrogen: 0.23,  pc: 1.7,  carbon: 0,    name: 'KOWEPO (서부발전)' },
    komipo:  { sulfur: 3.27, ash: 0.03,  nitrogen: 0.39,  pc: 1.27, carbon: 0,    name: 'KOMIPO (중부발전)' },
    ewp:     { sulfur: 3.81, ash: 0.034, nitrogen: 0.412, pc: 3.8,  carbon: 7.04, name: 'EWP (동서발전)' },
    koenipp: { sulfur: 4.53, ash: 0.09,  nitrogen: 0.39,  pc: 1.88, carbon: 0,    name: 'KOEN IPP' },
    gsg:     { sulfur: 2.96, ash: 0.1,   nitrogen: 0.21,  pc: 2.7,  carbon: 0,    name: 'GSG' },
    posco:   { sulfur: 4.26, ash: 0.07,  nitrogen: 0.37,  pc: 2.09, carbon: 0,    name: 'POSCO' },
    custom:  { sulfur: 3.00, ash: 0.07,  nitrogen: 0.30,  pc: 1.50, carbon: 0,    name: '사용자 정의 계수' }
};

// 기본 광산 데이터베이스
// baseFreight: ICI 기준 운임 8.47 기준의 각 광산 운임
// freightDiff: 기준 대비 차이 (freight - 8.47)
const DEFAULT_MINE_DATABASE = [
    { id: '1', name: '아다로', supplier: 'Adaro', nar: 4450, sulfur: 0.13, ash: 3.5, nitrogen: 1.1, baseFreight: 9.25, freightDiff: 0.78 },
    { id: '2', name: 'KPC', supplier: 'Kaltim Prima Coal', nar: 4350, sulfur: 0.45, ash: 7.5, nitrogen: 1.5, baseFreight: 8.24, freightDiff: -0.23 },
    { id: '3', name: '아르투민(55)', supplier: 'Arutmin', nar: 5650, sulfur: 0.55, ash: 9.5, nitrogen: 1.5, baseFreight: 10, freightDiff: 1.53 },
    { id: '4', name: '아르투민(47)', supplier: 'Arutmin', nar: 4850, sulfur: 0.4, ash: 8, nitrogen: 1.2, baseFreight: 10, freightDiff: 1.53 },
    { id: '5', name: 'MA', supplier: 'MA', nar: 5400, sulfur: 0.12, ash: 4.5, nitrogen: 1.5, baseFreight: 8.15, freightDiff: -0.32 },
    { id: '6', name: '부킷', supplier: 'Bukit', nar: 4600, sulfur: 0.5, ash: 7, nitrogen: 1.2, baseFreight: 9.09, freightDiff: 0.62 },
    { id: '7', name: '룸비', supplier: 'Lumbi', nar: 5000, sulfur: 0.45, ash: 15, nitrogen: 2.2, baseFreight: 12.56, freightDiff: 4.09 },
    { id: '8', name: '디자', supplier: 'Dija', nar: 4300, sulfur: 0.2, ash: 5, nitrogen: 1.2, baseFreight: 9.28, freightDiff: 0.81 },
    { id: '9', name: '롤디', supplier: 'Roldi', nar: 5400, sulfur: 0.5, ash: 12, nitrogen: 1.5, baseFreight: 12.77, freightDiff: 4.3 },
    { id: '10', name: '아르투민(42)', supplier: 'Arutmin', nar: 4350, sulfur: 0.25, ash: 6, nitrogen: 1.2, baseFreight: 10, freightDiff: 1.53 },
    { id: '11', name: '남아공탄', supplier: 'South Africa', nar: 5800, sulfur: 0.6, ash: 16, nitrogen: 2, baseFreight: 19, freightDiff: 10.53 },
    { id: '12', name: '제라', supplier: 'Jera', nar: 5725, sulfur: 0.6, ash: 17, nitrogen: 1.4, baseFreight: 14.04, freightDiff: 5.57 },
    { id: '13', name: '콜롬비아', supplier: 'Colombia', nar: 5800, sulfur: 0.6, ash: 12, nitrogen: 2, baseFreight: 30, freightDiff: 21.53 },
    { id: '14', name: 'ABB', supplier: 'ABB', nar: 5600, sulfur: 0.6, ash: 10, nitrogen: 1.7, baseFreight: 9.27, freightDiff: 0.8 },
    { id: '15', name: '바우', supplier: 'Bau', nar: 4750, sulfur: 0.5, ash: 8, nitrogen: 1.2, baseFreight: 9, freightDiff: 0.53 },
    { id: '16', name: 'KPUC', supplier: 'KPUC', nar: 5150, sulfur: 0.12, ash: 4.5, nitrogen: 1.5, baseFreight: 8.4, freightDiff: -0.07 },
    { id: '17', name: '엔탁', supplier: 'Entak', nar: 4875, sulfur: 0.35, ash: 6, nitrogen: 1.1, baseFreight: 14.25, freightDiff: 5.78 }
];

// 기준 운임 (이 값이 변경되면 모든 광산 운임이 연동됨)
const BASE_ICI_FREIGHT = 8.47;

// 애플리케이션 상태
let state = {
    mineDatabase: [],
    evaluationRows: [],
    rowCounter: 0,
    currentDraftId: null,
    iciEvaluation: 0  // ICI-5000 기준 Evaluation
};

// ==================== 로그인 및 초기화 ====================

async function handleLogin() {
    const email = document.getElementById('loginEmail').value.trim().toLowerCase();
    if (!email || !email.includes('@')) {
        showToast('유효한 이메일 주소를 입력하세요', 'error');
        return;
    }
    
    // 检查用户是否被授权
    if (!isAuthorized(email)) {
        showToast('권한 없음: 접속이 허용된 사용자가 아닙니다', 'error');
        return;
    }
    
    const userRole = getUserRole(email);
    const isAdmin = userRole === 'admin';
    
    localStorage.setItem('coal_user_email', email);
    localStorage.setItem('coal_user_role', userRole);
    
    await api.initialize(email);
    
    document.getElementById('loginOverlay').classList.add('hidden');
    updateUserDisplay(email, isAdmin, userRole);
    await initializeApp();
}

function updateUserDisplay(email, isAdmin, role) {
    const displayName = getUserName(email);
    document.getElementById('userEmail').textContent = displayName;
    
    if (isAdmin) {
        document.getElementById('adminTag').style.display = 'inline';
        document.querySelectorAll('.admin-only').forEach(el => {
            el.classList.add('visible');
        });
    }
    
    // 모든 사용자에게 개인 초안 기능 제공
    document.querySelectorAll('.user-only').forEach(el => {
        el.classList.add('visible');
    });
    
    const syncStatus = document.getElementById('syncStatus');
    if (api.isCloudEnabled()) {
        syncStatus.innerHTML = '<span class="sync-dot online"></span><span>클라우드 연결됨</span>';
    } else {
        syncStatus.innerHTML = '<span class="sync-dot offline"></span><span>로컬 모드</span>';
    }
}

document.addEventListener('DOMContentLoaded', function() {
    // Check if user is already logged in
    const savedEmail = localStorage.getItem('coal_user_email');
    if (savedEmail && isAuthorized(savedEmail)) {
        document.getElementById('loginEmail').value = savedEmail;
        handleLogin();
    }
});

// 로그아웃 기능
function handleLogout() {
    if (confirm('로그아웃하시겠습니까?')) {
        localStorage.removeItem('coal_user_email');
        localStorage.removeItem('coal_user_role');
        location.reload();
    }
}

async function initializeApp() {
    updateDateDisplay();
    await loadMineDatabase();
    await loadICIBaseline();
    renderMineDatabase();
    bindEventListeners();
    calculateICIEvaluation();
    
    // 초기화 시 빈 행 추가하지 않음
    if (state.evaluationRows.length > 0) {
        renderAllRows();
    }
}

function updateDateDisplay() {
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
    document.getElementById('currentDate').textContent = now.toLocaleDateString('ko-KR', options);
}

async function loadMineDatabase() {
    const result = await api.getMineDatabase();
    if (result.success && result.data && result.data.length > 0) {
        state.mineDatabase = result.data;
    } else {
        state.mineDatabase = [...DEFAULT_MINE_DATABASE];
    }
}

async function loadICIBaseline() {
    const result = await api.getICIBaseline();
    if (result.success && result.data) {
        if (result.data.fobt) document.getElementById('iciFOBT').value = result.data.fobt;
        if (result.data.freight) document.getElementById('iciFreight').value = result.data.freight;
        if (result.data.sulfur) document.getElementById('iciSulfur').value = result.data.sulfur;
        if (result.data.ash) document.getElementById('iciAsh').value = result.data.ash;
        if (result.data.nitrogen) document.getElementById('iciNitrogen').value = result.data.nitrogen;
        if (result.data.nar) document.getElementById('iciNAR').value = result.data.nar;
        if (result.data.exchangeRate) document.getElementById('exchangeRate').value = result.data.exchangeRate;
        if (result.data.consumptionTaxKRW) document.getElementById('consumptionTaxKRW').value = result.data.consumptionTaxKRW;
    }
}

function bindEventListeners() {
    // ICI 파라미터 변경
    ['iciFOBT', 'iciSulfur', 'iciAsh', 'iciNitrogen', 'iciNAR', 'exchangeRate', 'consumptionTaxKRW'].forEach(id => {
        document.getElementById(id).addEventListener('input', () => {
            calculateICIEvaluation();
            recalculateAll();
        });
    });
    
    // ICI 기준 운임 변경 시 모든 광산 운임 업데이트
    document.getElementById('iciFreight').addEventListener('input', () => {
        updateAllFreights();
        calculateICIEvaluation();
        recalculateAll();
    });
    
    // 발전사 선택
    document.getElementById('gencoSelect').addEventListener('change', onGencoChange);
    
    // 페널티 계수 변경
    ['coefSulfur', 'coefAsh', 'coefNitrogen', 'coefPC', 'coefCarbon'].forEach(id => {
        document.getElementById(id).addEventListener('input', () => {
            calculateICIEvaluation();
            recalculateAll();
        });
    });
}

// ==================== 운임 연동 업데이트 ====================

function updateAllFreights() {
    const currentICIFreight = parseFloat(document.getElementById('iciFreight').value) || BASE_ICI_FREIGHT;
    
    state.evaluationRows.forEach(row => {
        // freightDiff가 있는 경우에만 자동 업데이트 (광산 DB에서 추가된 경우)
        if (row.freightDiff !== undefined && row.freightDiff !== null) {
            const newFreight = currentICIFreight + row.freightDiff;
            row.freight = newFreight.toFixed(2);
            
            // UI 업데이트
            const tr = document.getElementById(`row-${row.id}`);
            if (tr) {
                const freightInput = tr.querySelector('input[onchange*="freight"]');
                if (freightInput) {
                    freightInput.value = newFreight.toFixed(2);
                }
            }
        }
    });
}

// ==================== ICI 기준 Evaluation 계산 ====================

function calculateICIEvaluation() {
    const fobt = parseFloat(document.getElementById('iciFOBT').value) || 0;
    const freight = parseFloat(document.getElementById('iciFreight').value) || 0;
    const sulfur = parseFloat(document.getElementById('iciSulfur').value) || 0;
    const ash = parseFloat(document.getElementById('iciAsh').value) || 0;
    const nitrogen = parseFloat(document.getElementById('iciNitrogen').value) || 0;
    const nar = parseFloat(document.getElementById('iciNAR').value) || 5000;
    
    const coefSulfur = parseFloat(document.getElementById('coefSulfur').value) || 0;
    const coefAsh = parseFloat(document.getElementById('coefAsh').value) || 0;
    const coefNitrogen = parseFloat(document.getElementById('coefNitrogen').value) || 0;
    const coefPC = parseFloat(document.getElementById('coefPC').value) || 0;
    
    const taxKRW = parseFloat(document.getElementById('consumptionTaxKRW').value) || 0;
    const rate = parseFloat(document.getElementById('exchangeRate').value) || 1450;
    const consumptionTaxUSD = taxKRW / rate;
    
    // CFR
    const cfr = fobt + freight;
    
    // Quality Penalty
    const penalty = (sulfur * coefSulfur) + (ash * coefAsh) + (nitrogen * coefNitrogen);
    
    // Evaluation = (CFR + Penalty + P/C + Consumption Tax) × (6080 / NAR)
    const evaluation = (cfr + penalty + coefPC + consumptionTaxUSD) * (6080 / nar);
    
    // 상태 저장
    state.iciEvaluation = evaluation;
    
    // UI 업데이트
    document.getElementById('consumptionTaxUSD').textContent = consumptionTaxUSD.toFixed(2);
    document.getElementById('iciEvaluation').textContent = evaluation.toFixed(2);
    
    // 安全检查：如果 HTML 中有这个元素才更新
    const baselineEvalEl = document.getElementById('baselineEval');
    if (baselineEvalEl) {
        baselineEvalEl.textContent = evaluation.toFixed(2);
    }
    
    return evaluation;
}

// ==================== 토스트 메시지 ====================

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

// ==================== ICI 기준 ====================

async function syncICIFromCloud() {
    showToast('ICI 기준 동기화 중...', 'info');
    await loadICIBaseline();
    calculateICIEvaluation();
    recalculateAll();
    showToast('ICI 기준이 동기화되었습니다', 'success');
}

async function saveICIToCloud() {
    if (!api.canEditICI()) {
        showToast('권한 부족: 관리자만 ICI 기준을 수정할 수 있습니다', 'error');
        return;
    }
    
    const data = {
        fobt: document.getElementById('iciFOBT').value,
        freight: document.getElementById('iciFreight').value,
        sulfur: document.getElementById('iciSulfur').value,
        ash: document.getElementById('iciAsh').value,
        nitrogen: document.getElementById('iciNitrogen').value,
        nar: document.getElementById('iciNAR').value,
        exchangeRate: document.getElementById('exchangeRate').value,
        consumptionTaxKRW: document.getElementById('consumptionTaxKRW').value
    };
    
    const result = await api.updateICIBaseline(data);
    if (result.success) {
        showToast('ICI 기준이 클라우드에 저장되었습니다', 'success');
    } else {
        showToast(result.error || '저장 실패', 'error');
    }
}

// ==================== 발전사 관리 ====================

function onGencoChange() {
    const genco = document.getElementById('gencoSelect').value;
    const coefficients = GENCO_COEFFICIENTS[genco];
    
    if (coefficients && genco !== 'custom') {
        document.getElementById('coefSulfur').value = coefficients.sulfur;
        document.getElementById('coefAsh').value = coefficients.ash;
        document.getElementById('coefNitrogen').value = coefficients.nitrogen;
        document.getElementById('coefPC').value = coefficients.pc;
        document.getElementById('coefCarbon').value = coefficients.carbon;
    }
    
    // ICI 기준 Evaluation 재계산
    calculateICIEvaluation();
    
    // 모든 행 재계산
    state.evaluationRows.forEach(row => {
        calculateRow(row.id);
    });
    
    // 하이라이트 및 요약 업데이트
    highlightBestAndWorst();
    updateSummary();
}

// ==================== 평가표 작업 ====================

function addMineRow(prefillData = null) {
    state.rowCounter++;
    const rowId = state.rowCounter;
    
    const row = {
        id: rowId,
        mineName: prefillData?.name || '',
        fobt: prefillData?.fobt || '',
        freight: '',
        sulfur: prefillData?.sulfur || '',
        ash: prefillData?.ash || '',
        nitrogen: prefillData?.nitrogen || '',
        nar: prefillData?.nar || '',
        cfr: null,
        penalty: null,
        evaluation: null,
        delta: null  // 평가상 (차이)
    };
    
    state.evaluationRows.push(row);
    renderRow(row);
    updateSummary();
}

function renderRow(row) {
    const tbody = document.getElementById('evalTableBody');
    const tr = document.createElement('tr');
    tr.id = `row-${row.id}`;
    tr.dataset.rowId = row.id;
    
    const mineOptions = state.mineDatabase.map(m => 
        `<option value="${m.id}">${m.name}</option>`
    ).join('');
    
    // 如果有名字，显示名字输入框；否则显示下拉选择
    const hasName = row.mineName && row.mineName.trim() !== '';
    const showNameInput = hasName ? 'block' : 'none';
    const showSelect = hasName ? 'none' : 'block';
    
    tr.innerHTML = `
        <td>
            <button class="btn-delete-row" onclick="deleteRow(${row.id})">×</button>
        </td>
        <td>
            <select class="mine-select" style="display:${showSelect};" onchange="onMineSelect(${row.id}, this.value)">
                <option value="">Bidder 선택...</option>
                ${mineOptions}
                <option value="custom">직접 입력</option>
            </select>
            <input type="text" class="mine-name-input" style="display:${showNameInput};" 
                   placeholder="Bidder명" value="${row.mineName}" 
                   onchange="updateRowData(${row.id}, 'mineName', this.value)">
        </td>
        <td>
            <input type="number" step="0.01" value="${row.fobt}" 
                   onchange="updateRowData(${row.id}, 'fobt', this.value)">
        </td>
        <td>
            <input type="number" step="0.01" value="${row.freight}" 
                   onchange="updateRowData(${row.id}, 'freight', this.value)">
        </td>
        <td>
            <input type="number" step="0.01" value="${row.sulfur}" 
                   onchange="updateRowData(${row.id}, 'sulfur', this.value)">
        </td>
        <td>
            <input type="number" step="0.01" value="${row.ash}" 
                   onchange="updateRowData(${row.id}, 'ash', this.value)">
        </td>
        <td>
            <input type="number" step="0.01" value="${row.nitrogen}" 
                   onchange="updateRowData(${row.id}, 'nitrogen', this.value)">
        </td>
        <td>
            <input type="number" step="10" value="${row.nar}" 
                   onchange="updateRowData(${row.id}, 'nar', this.value)">
        </td>
        <td class="result-cell cfr-result">--</td>
        <td class="result-cell penalty-result">--</td>
        <td class="result-cell eval-result">--</td>
        <td class="result-cell delta-result">--</td>
    `;
    
    tbody.appendChild(tr);
}

function renderAllRows() {
    const tbody = document.getElementById('evalTableBody');
    tbody.innerHTML = '';
    state.evaluationRows.forEach(row => renderRow(row));
    recalculateAll();
}

function onMineSelect(rowId, mineId) {
    const row = state.evaluationRows.find(r => r.id === rowId);
    const tr = document.getElementById(`row-${rowId}`);
    const nameInput = tr.querySelector('.mine-name-input');
    
    if (mineId === 'custom') {
        nameInput.style.display = 'block';
        row.mineName = '';
        row.mineId = null;
        row.freightDiff = 0;
    } else if (mineId) {
        const mine = state.mineDatabase.find(m => m.id === mineId || m.id === parseInt(mineId));
        if (mine) {
            // 현재 ICI 기준 운임
            const currentICIFreight = parseFloat(document.getElementById('iciFreight').value) || BASE_ICI_FREIGHT;
            const calculatedFreight = currentICIFreight + (mine.freightDiff || 0);
            
            nameInput.style.display = 'none';
            row.mineName = mine.name;
            row.mineId = mine.id;
            row.freightDiff = mine.freightDiff || 0;
            row.nar = mine.nar;
            row.sulfur = mine.sulfur;
            row.ash = mine.ash;
            row.nitrogen = mine.nitrogen;
            row.freight = calculatedFreight.toFixed(2);
            
            tr.querySelector('input[onchange*="nar"]').value = mine.nar;
            tr.querySelector('input[onchange*="sulfur"]').value = mine.sulfur;
            tr.querySelector('input[onchange*="ash"]').value = mine.ash;
            tr.querySelector('input[onchange*="nitrogen"]').value = mine.nitrogen;
            tr.querySelector('input[onchange*="freight"]').value = calculatedFreight.toFixed(2);
            
            calculateRow(rowId);
        }
    } else {
        nameInput.style.display = 'none';
    }
}

function updateRowData(rowId, field, value) {
    const row = state.evaluationRows.find(r => r.id === rowId);
    if (row) {
        row[field] = value;
        calculateRow(rowId);
    }
}

/**
 * 핵심 계산 공식:
 * 
 * Evaluation = (CFR + Quality Penalty + P/C + Consumption Tax) × (6080 / NAR)
 * 평가상 = Bidder Evaluation - ICI-5000 기준 Evaluation
 * 
 * CFR = FOBT + Freight
 * Quality Penalty = (Sulfur × 계수) + (Ash × 계수) + (Nitrogen × 계수)
 * P/C = Power Constant (발전사별 고정값)
 * Consumption Tax = 한화 ÷ 환율
 */
function calculateRow(rowId) {
    const row = state.evaluationRows.find(r => r.id === rowId);
    if (!row) return;
    
    const fobt = parseFloat(row.fobt) || 0;
    const freight = parseFloat(row.freight) || 0;
    const sulfur = parseFloat(row.sulfur) || 0;
    const ash = parseFloat(row.ash) || 0;
    const nitrogen = parseFloat(row.nitrogen) || 0;
    const nar = parseFloat(row.nar) || 0;
    
    // 계수 가져오기
    const coefSulfur = parseFloat(document.getElementById('coefSulfur').value) || 0;
    const coefAsh = parseFloat(document.getElementById('coefAsh').value) || 0;
    const coefNitrogen = parseFloat(document.getElementById('coefNitrogen').value) || 0;
    const coefPC = parseFloat(document.getElementById('coefPC').value) || 0;
    
    // Consumption Tax (USD)
    const taxKRW = parseFloat(document.getElementById('consumptionTaxKRW').value) || 0;
    const rate = parseFloat(document.getElementById('exchangeRate').value) || 1450;
    const consumptionTaxUSD = taxKRW / rate;
    
    // CFR = FOBT + Freight
    row.cfr = fobt + freight;
    
    // Quality Penalty = Sulfur×계수 + Ash×계수 + Nitrogen×계수
    row.penalty = (sulfur * coefSulfur) + (ash * coefAsh) + (nitrogen * coefNitrogen);
    
    // Evaluation = (CFR + Penalty + P/C + Consumption Tax) × (6080 / NAR)
    if (nar > 0) {
        row.evaluation = (row.cfr + row.penalty + coefPC + consumptionTaxUSD) * (6080 / nar);
    } else {
        row.evaluation = null;
    }
    
    // 평가상 = Bidder Evaluation - ICI-5000 기준 Evaluation
    if (row.evaluation !== null) {
        row.delta = row.evaluation - state.iciEvaluation;
    } else {
        row.delta = null;
    }
    
    updateRowDisplay(rowId);
    updateSummary();
}

function updateRowDisplay(rowId) {
    const row = state.evaluationRows.find(r => r.id === rowId);
    const tr = document.getElementById(`row-${rowId}`);
    if (!row || !tr) return;
    
    tr.querySelector('.cfr-result').textContent = row.cfr !== null ? row.cfr.toFixed(2) : '--';
    tr.querySelector('.penalty-result').textContent = row.penalty !== null ? row.penalty.toFixed(2) : '--';
    tr.querySelector('.eval-result').textContent = row.evaluation !== null ? row.evaluation.toFixed(2) : '--';
    
    const deltaCell = tr.querySelector('.delta-result');
    if (row.delta !== null) {
        deltaCell.textContent = row.delta.toFixed(2);
        // 낮을수록 좋음 (음수 = 녹색, 양수 = 빨간색)
        deltaCell.className = 'result-cell delta-result ' + (row.delta <= 0 ? 'delta-negative' : 'delta-positive');
    } else {
        deltaCell.textContent = '--';
        deltaCell.className = 'result-cell delta-result';
    }
}

function recalculateAll() {
    calculateICIEvaluation();
    state.evaluationRows.forEach(row => calculateRow(row.id));
    highlightBestAndWorst();
    updateSummary();
}

function highlightBestAndWorst() {
    const validRows = state.evaluationRows.filter(r => r.delta !== null);
    if (validRows.length === 0) return;
    
    // 가장 낮은 평가상 (가장 좋은 입찰)
    const minDelta = Math.min(...validRows.map(r => r.delta));
    
    document.querySelectorAll('#evalTableBody tr').forEach(tr => {
        tr.classList.remove('row-best', 'row-high');
    });
    
    validRows.forEach(row => {
        const tr = document.getElementById(`row-${row.id}`);
        if (!tr) return;
        
        if (row.delta === minDelta) {
            tr.classList.add('row-best');
        } else if (row.delta > 10) {
            tr.classList.add('row-high');
        }
    });
}

function deleteRow(rowId) {
    state.evaluationRows = state.evaluationRows.filter(r => r.id !== rowId);
    const tr = document.getElementById(`row-${rowId}`);
    if (tr) tr.remove();
    updateSummary();
    highlightBestAndWorst();
}

function clearAllRows() {
    if (confirm('모든 평가 데이터를 삭제하시겠습니까?')) {
        state.evaluationRows = [];
        state.currentDraftId = null;
        document.getElementById('evalTableBody').innerHTML = '';
        document.getElementById('batchName').value = '';
        updateSummary();
        addMineRow();
    }
}

function updateSummary() {
    const validRows = state.evaluationRows.filter(r => r.delta !== null);
    
    document.getElementById('bidCount').textContent = validRows.length;
    
    if (validRows.length === 0) {
        document.getElementById('bestBid').textContent = '--';
        document.getElementById('bestDetail').textContent = '--';
        document.getElementById('bestDeltaValue').textContent = '--';
        return;
    }
    
    // 가장 낮은 평가상 찾기
    const best = validRows.reduce((a, b) => a.delta < b.delta ? a : b);
    document.getElementById('bestBid').textContent = best.mineName || '이름 없음';
    document.getElementById('bestDetail').textContent = `평가상: ${best.delta.toFixed(2)} | Evaluation: ${best.evaluation.toFixed(2)}`;
    document.getElementById('bestDeltaValue').textContent = best.delta.toFixed(2);
}

// Debug function removed

// ==================== 광산 데이터베이스 ====================

function renderMineDatabase() {
    const container = document.getElementById('mineDBList');
    
    if (state.mineDatabase.length === 0) {
        container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">🏭</div><div>광산 데이터 없음</div></div>';
        return;
    }
    
    container.innerHTML = state.mineDatabase.map(mine => `
        <div class="mine-db-item" onclick="quickAddMine('${mine.id}')">
            <div class="mine-name">${mine.name}</div>
            <div class="mine-info">NAR ${mine.nar} | S ${mine.sulfur}% | A ${mine.ash}%</div>
        </div>
    `).join('');
}

function quickAddMine(mineId) {
    const mine = state.mineDatabase.find(m => m.id === mineId || m.id === parseInt(mineId));
    if (mine) {
        // 현재 ICI 기준 운임
        const currentICIFreight = parseFloat(document.getElementById('iciFreight').value) || BASE_ICI_FREIGHT;
        
        // 광산 운임 계산: 기준 운임 + 차이값
        const calculatedFreight = currentICIFreight + (mine.freightDiff || 0);
        
        // 添加行并预填数据
        state.rowCounter++;
        const rowId = state.rowCounter;
        
        const row = {
            id: rowId,
            mineName: mine.name,
            mineId: mine.id,
            fobt: '',
            freight: calculatedFreight.toFixed(2),
            freightDiff: mine.freightDiff || 0,
            sulfur: mine.sulfur,
            ash: mine.ash,
            nitrogen: mine.nitrogen,
            nar: mine.nar,
            cfr: null,
            penalty: null,
            evaluation: null,
            delta: null
        };
        
        state.evaluationRows.push(row);
        renderRow(row);
        calculateRow(rowId);
        updateSummary();
        
        showToast(`${mine.name} 추가됨`, 'success');
    }
}

function showAddMineDBModal() {
    document.getElementById('addMineDBModal').classList.add('active');
}

async function addMineToDatabase() {
    const name = document.getElementById('newMineName').value.trim();
    const supplier = document.getElementById('newMineSupplier').value.trim();
    const nar = parseFloat(document.getElementById('newMineNAR').value) || 0;
    const sulfur = parseFloat(document.getElementById('newMineSulfur').value) || 0;
    const ash = parseFloat(document.getElementById('newMineAsh').value) || 0;
    const nitrogen = parseFloat(document.getElementById('newMineNitrogen').value) || 0;
    
    if (!name) {
        showToast('광산명을 입력하세요', 'error');
        return;
    }
    
    const result = await api.addMine({ name, supplier, nar, sulfur, ash, nitrogen });
    
    if (result.success) {
        state.mineDatabase.push({ id: result.data.id, name, supplier, nar, sulfur, ash, nitrogen });
        renderMineDatabase();
        closeModal('addMineDBModal');
        showToast(`광산 "${name}"이(가) 추가되었습니다`, 'success');
        
        document.getElementById('newMineName').value = '';
        document.getElementById('newMineSupplier').value = '';
        document.getElementById('newMineNAR').value = '4200';
        document.getElementById('newMineSulfur').value = '0.10';
        document.getElementById('newMineAsh').value = '5.00';
        document.getElementById('newMineNitrogen').value = '0.80';
    } else {
        showToast(result.error || '추가 실패', 'error');
    }
}

// ==================== 초안 관리 ====================

function getCurrentBidData() {
    const validRows = state.evaluationRows.filter(r => r.delta !== null);
    const best = validRows.length > 0 
        ? validRows.reduce((a, b) => a.delta < b.delta ? a : b) 
        : null;
    
    return {
        iciParams: {
            fobt: document.getElementById('iciFOBT').value,
            freight: document.getElementById('iciFreight').value,
            sulfur: document.getElementById('iciSulfur').value,
            ash: document.getElementById('iciAsh').value,
            nitrogen: document.getElementById('iciNitrogen').value,
            nar: document.getElementById('iciNAR').value,
            exchangeRate: document.getElementById('exchangeRate').value,
            consumptionTaxKRW: document.getElementById('consumptionTaxKRW').value
        },
        genco: document.getElementById('gencoSelect').value,
        coefficients: {
            sulfur: document.getElementById('coefSulfur').value,
            ash: document.getElementById('coefAsh').value,
            nitrogen: document.getElementById('coefNitrogen').value,
            pc: document.getElementById('coefPC').value,
            carbon: document.getElementById('coefCarbon').value
        },
        iciEvaluation: state.iciEvaluation,
        bestMine: best?.mineName || '',
        bestDelta: best?.delta?.toFixed(2) || '',
        rows: state.evaluationRows.map(r => ({...r}))
    };
}

function loadBidData(data) {
    // ICI 파라미터 복원
    if (data.iciParams.fobt) document.getElementById('iciFOBT').value = data.iciParams.fobt;
    if (data.iciParams.freight) document.getElementById('iciFreight').value = data.iciParams.freight;
    if (data.iciParams.sulfur) document.getElementById('iciSulfur').value = data.iciParams.sulfur;
    if (data.iciParams.ash) document.getElementById('iciAsh').value = data.iciParams.ash;
    if (data.iciParams.nitrogen) document.getElementById('iciNitrogen').value = data.iciParams.nitrogen;
    if (data.iciParams.nar) document.getElementById('iciNAR').value = data.iciParams.nar;
    if (data.iciParams.exchangeRate) document.getElementById('exchangeRate').value = data.iciParams.exchangeRate;
    if (data.iciParams.consumptionTaxKRW) document.getElementById('consumptionTaxKRW').value = data.iciParams.consumptionTaxKRW;
    
    // 발전사 및 계수 복원
    document.getElementById('gencoSelect').value = data.genco;
    document.getElementById('coefSulfur').value = data.coefficients.sulfur;
    document.getElementById('coefAsh').value = data.coefficients.ash;
    document.getElementById('coefNitrogen').value = data.coefficients.nitrogen;
    if (data.coefficients.pc) document.getElementById('coefPC').value = data.coefficients.pc;
    if (data.coefficients.carbon) document.getElementById('coefCarbon').value = data.coefficients.carbon;
    
    // 평가 행 복원
    state.evaluationRows = data.rows.map(r => ({...r}));
    state.rowCounter = Math.max(...state.evaluationRows.map(r => r.id), 0);
    
    calculateICIEvaluation();
    renderAllRows();
}

async function saveDraft() {
    const name = document.getElementById('batchName').value.trim();
    if (!name) {
        showToast('초안 이름을 입력하세요', 'warning');
        document.getElementById('batchName').focus();
        return;
    }
    
    const data = {
        id: state.currentDraftId,
        name: name,
        fullData: getCurrentBidData()
    };
    
    const result = await api.saveDraft(data);
    
    if (result.success) {
        state.currentDraftId = result.data.id;
        showToast(`초안 "${name}"이(가) 저장되었습니다`, 'success');
    } else {
        showToast(result.error || '저장 실패', 'error');
    }
}

async function showDraftsModal() {
    const result = await api.getMyDrafts();
    const container = document.getElementById('draftsList');
    
    if (!result.success || !result.data || result.data.length === 0) {
        container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">📁</div><div>저장된 초안이 없습니다</div></div>';
    } else {
        container.innerHTML = result.data.map(draft => `
            <div class="batch-item">
                <div class="batch-item-info">
                    <div class="batch-name">${draft.name}</div>
                    <div class="batch-date">${new Date(draft.updatedAt || draft.createdAt).toLocaleString('ko-KR')}</div>
                </div>
                <div class="batch-item-actions">
                    <button class="btn btn-primary btn-small" onclick="loadDraft('${draft.id}', '${draft.name}')">불러오기</button>
                    <button class="btn btn-danger btn-small" onclick="deleteDraft('${draft.id}')">삭제</button>
                </div>
            </div>
        `).join('');
    }
    
    document.getElementById('draftsModal').classList.add('active');
}

async function loadDraft(draftId, draftName) {
    const result = await api.getMyDrafts();
    const draft = result.data?.find(d => d.id === draftId);
    
    if (draft && draft.fullData) {
        state.currentDraftId = draftId;
        document.getElementById('batchName').value = draft.name;
        loadBidData(draft.fullData);
        closeModal('draftsModal');
        showToast(`초안 불러옴: ${draftName}`, 'success');
    } else {
        showToast('불러오기 실패', 'error');
    }
}

async function deleteDraft(draftId) {
    if (!confirm('이 초안을 삭제하시겠습니까?')) return;
    
    const result = await api.deleteDraft(draftId);
    if (result.success) {
        if (state.currentDraftId === draftId) {
            state.currentDraftId = null;
        }
        showToast('초안이 삭제되었습니다', 'success');
        showDraftsModal();
    } else {
        showToast(result.error || '삭제 실패', 'error');
    }
}

// ==================== 공식 입찰 ====================

async function saveOfficialBid() {
    if (!api.canCreateOfficialBid()) {
        showToast('권한 부족: 관리자만 공식 입찰 평가를 생성할 수 있습니다', 'error');
        return;
    }
    
    const name = document.getElementById('batchName').value.trim();
    if (!name) {
        showToast('입찰 배치 이름을 입력하세요', 'warning');
        document.getElementById('batchName').focus();
        return;
    }
    
    const bidData = getCurrentBidData();
    
    const data = {
        name: name,
        genco: bidData.genco,
        bestMine: bidData.bestMine,
        bestDelta: bidData.bestDelta,
        fullData: bidData,
        status: 'final'
    };
    
    const result = await api.createOfficialBid(data);
    
    if (result.success) {
        showToast(`공식 입찰 평가 "${name}"이(가) 생성되었습니다`, 'success');
    } else {
        showToast(result.error || '생성 실패', 'error');
    }
}

async function showOfficialBidsModal() {
    const result = await api.getOfficialBids();
    const container = document.getElementById('officialBidsList');
    
    if (!result.success || !result.data || result.data.length === 0) {
        container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">📋</div><div>공식 입찰 평가 기록이 없습니다</div></div>';
    } else {
        container.innerHTML = result.data.map(bid => `
            <div class="batch-item">
                <div class="batch-item-info">
                    <div class="batch-name">${bid.name} <span class="status-tag ${bid.status}">${bid.status === 'final' ? '공식' : '초안'}</span></div>
                    <div class="batch-date">
                        ${new Date(bid.createdAt).toLocaleString('ko-KR')} | 
                        최적: ${bid.bestMine || '--'} (평가상: ${bid.bestDelta || bid.bestEvalValue || '--'}) |
                        생성자: ${bid.createdBy || '--'}
                    </div>
                </div>
                <div class="batch-item-actions">
                    <button class="btn btn-primary btn-small" onclick="loadOfficialBid('${bid.id}', '${bid.name}')">보기</button>
                </div>
            </div>
        `).join('');
    }
    
    document.getElementById('officialBidsModal').classList.add('active');
}

async function loadOfficialBid(bidId, bidName) {
    const result = await api.getOfficialBids();
    const bid = result.data?.find(b => b.id === bidId);
    
    if (bid && bid.fullData) {
        document.getElementById('batchName').value = bid.name;
        loadBidData(bid.fullData);
        closeModal('officialBidsModal');
        showToast(`공식 입찰 평가 불러옴: ${bidName}`, 'success');
    } else {
        showToast('불러오기 실패', 'error');
    }
}

// ==================== 배치 비교 ====================

async function showCompareModal() {
    const draftsResult = await api.getMyDrafts();
    const officialResult = await api.getOfficialBids();
    
    const allBatches = [
        ...(draftsResult.data || []).map(d => ({ ...d, type: 'draft' })),
        ...(officialResult.data || []).map(b => ({ ...b, type: 'official' }))
    ];
    
    const selectA = document.getElementById('compareBatchA');
    const selectB = document.getElementById('compareBatchB');
    
    const options = allBatches.map(b => 
        `<option value="${b.id}" data-type="${b.type}">${b.type === 'official' ? '📋' : '📝'} ${b.name}</option>`
    ).join('');
    
    selectA.innerHTML = '<option value="">배치 선택...</option>' + options;
    selectB.innerHTML = '<option value="">배치 선택...</option>' + options;
    
    document.getElementById('compareResults').innerHTML = '';
    document.getElementById('compareModal').classList.add('active');
}

async function runComparison() {
    const batchIdA = document.getElementById('compareBatchA').value;
    const batchIdB = document.getElementById('compareBatchB').value;
    
    if (!batchIdA || !batchIdB) {
        showToast('비교할 두 배치를 선택하세요', 'warning');
        return;
    }
    
    const draftsResult = await api.getMyDrafts();
    const officialResult = await api.getOfficialBids();
    
    const allBatches = [...(draftsResult.data || []), ...(officialResult.data || [])];
    
    const batchA = allBatches.find(b => b.id === batchIdA);
    const batchB = allBatches.find(b => b.id === batchIdB);
    
    if (!batchA || !batchB) {
        showToast('배치 데이터가 불완전합니다', 'error');
        return;
    }
    
    const dataA = batchA.fullData;
    const dataB = batchB.fullData;
    
    const allMines = new Set([
        ...dataA.rows.map(r => r.mineName),
        ...dataB.rows.map(r => r.mineName)
    ]);
    
    let html = `
        <table>
            <thead>
                <tr>
                    <th>Bidder</th>
                    <th>${batchA.name}<br>FOBT</th>
                    <th>${batchB.name}<br>FOBT</th>
                    <th>FOBT 변화</th>
                    <th>${batchA.name}<br>평가상</th>
                    <th>${batchB.name}<br>평가상</th>
                    <th>평가상 변화</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    allMines.forEach(mineName => {
        if (!mineName) return;
        
        const rowA = dataA.rows.find(r => r.mineName === mineName);
        const rowB = dataB.rows.find(r => r.mineName === mineName);
        
        const fobtA = rowA?.fobt || '--';
        const fobtB = rowB?.fobt || '--';
        const deltaA = rowA?.delta?.toFixed(2) || '--';
        const deltaB = rowB?.delta?.toFixed(2) || '--';
        
        let fobtChange = '--';
        let fobtClass = '';
        if (rowA?.fobt && rowB?.fobt) {
            const diff = parseFloat(rowB.fobt) - parseFloat(rowA.fobt);
            fobtChange = (diff >= 0 ? '+' : '') + diff.toFixed(2);
            fobtClass = diff > 0 ? 'change-up' : (diff < 0 ? 'change-down' : '');
        }
        
        let deltaChange = '--';
        let deltaClass = '';
        if (rowA?.delta !== null && rowB?.delta !== null) {
            const diff = rowB.delta - rowA.delta;
            deltaChange = (diff >= 0 ? '+' : '') + diff.toFixed(2);
            deltaClass = diff > 0 ? 'change-up' : (diff < 0 ? 'change-down' : '');
        }
        
        html += `
            <tr>
                <td style="text-align:left;font-weight:500;">${mineName}</td>
                <td>${fobtA}</td>
                <td>${fobtB}</td>
                <td class="${fobtClass}">${fobtChange}</td>
                <td>${deltaA}</td>
                <td>${deltaB}</td>
                <td class="${deltaClass}">${deltaChange}</td>
            </tr>
        `;
    });
    
    html += '</tbody></table>';
    
    document.getElementById('compareResults').innerHTML = html;
}

// ==================== Excel导出 (精确还原共同入札图表布局) ====================
function exportToExcel() {
    const wb = XLSX.utils.book_new();
    
    // 严格按照图中的发电厂顺序：EWP, KOEN, KOWEPO, KOMIPO, KOSPO
    const targetGencos = ['ewp', 'koen', 'kowepo', 'komipo', 'kospo'];
    
    // 获取页面基础参数
    const fobt = parseFloat(document.getElementById('iciFOBT').value) || 0;
    const freight = parseFloat(document.getElementById('iciFreight').value) || 0;
    const sulfur = parseFloat(document.getElementById('iciSulfur').value) || 0;
    const ash = parseFloat(document.getElementById('iciAsh').value) || 0;
    const nitrogen = parseFloat(document.getElementById('iciNitrogen').value) || 0;
    const nar = parseFloat(document.getElementById('iciNAR').value) || 5000;
    const rate = document.getElementById('exchangeRate').value || '1450';
    const taxUSD = document.getElementById('consumptionTaxUSD').textContent || '31.72';
    const iciCfr = fobt + freight;
    
    // 预计算 ICI-5000 在各大发电厂的 Evaluation 基准值
    const iciEvaluations = {};
    targetGencos.forEach(gKey => {
        const coefs = GENCO_COEFFICIENTS[gKey];
        const penalty = (sulfur * coefs.sulfur) + (ash * coefs.ash) + (nitrogen * coefs.nitrogen);
        iciEvaluations[gKey] = (iciCfr + penalty + coefs.pc + parseFloat(taxUSD)) * (6080 / nar);
    });
    
    // 1. 构建顶部参数行 (与截图一致)
    const exportData = [
        ['공동입찰 종합결과'], // 行1: 大标题
        [], // 行2: 空行
        // 行3: 汇率与消费税
        ['', `적용환율 (원/$): ${Number(rate).toLocaleString()}`, '', '', '', '', `소비세 ($/mt): ${parseFloat(taxUSD).toFixed(2)}`]
    ];
    
    // 2. 构建复合表头 (行4 和 行5)
    const headerRow1 = ['No.', 'Bidder', 'Supplier', 'Mine', 'NAR\n(kcal/kg)', 'FOBT\n($/ton)', 'Freight\n($/ton)', 'CFR\n($/ton)'];
    const headerRow2 = ['', '', '', '', '', '', '', ''];
    targetGencos.forEach(gKey => {
        headerRow1.push(gKey.toUpperCase(), '');
        headerRow2.push('Evaluation', '평가상(Delta)');
    });
    exportData.push(headerRow1, headerRow2);
    
    // 3. 插入固定第一行数据: ICI-5000 (行6)
    const iciRow = [1, 'ICI-5000', 'ICI-5000', 'ICI-5000', nar, fobt.toFixed(2), freight.toFixed(2), iciCfr.toFixed(2)];
    targetGencos.forEach(gKey => {
        iciRow.push(iciEvaluations[gKey].toFixed(2), '0.00');
    });
    exportData.push(iciRow);
    
    // 4. 遍历插入所有 Bidder 的数据
    state.evaluationRows.forEach((row, index) => {
        const mineInfo = state.mineDatabase.find(m => m.id === row.mineId) || {};
        const supplier = mineInfo.supplier || row.mineName || '';
        const mineName = row.mineName || '';
        const rFobt = parseFloat(row.fobt) || 0;
        const rFreight = parseFloat(row.freight) || 0;
        const rowCfr = row.cfr !== null ? row.cfr : (rFobt + rFreight);
        
        const dataRow = [
            index + 2,
            mineName,
            supplier,
            mineName,
            row.nar || '',
            rFobt ? rFobt.toFixed(2) : '',
            rFreight ? rFreight.toFixed(2) : '',
            rowCfr ? rowCfr.toFixed(2) : ''
        ];
        
        targetGencos.forEach(gKey => {
            if (!row.nar || row.nar <= 0) {
                dataRow.push('--', '--');
                return;
            }
            const coefs = GENCO_COEFFICIENTS[gKey];
            const rSulfur = parseFloat(row.sulfur) || 0;
            const rAsh = parseFloat(row.ash) || 0;
            const rNitrogen = parseFloat(row.nitrogen) || 0;
            const penalty = (rSulfur * coefs.sulfur) + (rAsh * coefs.ash) + (rNitrogen * coefs.nitrogen);
            const evaluation = (rowCfr + penalty + coefs.pc + parseFloat(taxUSD)) * (6080 / row.nar);
            const delta = evaluation - iciEvaluations[gKey];
            dataRow.push(evaluation.toFixed(2), delta.toFixed(2));
        });
        
        exportData.push(dataRow);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(exportData);
    
    // 5. 单元格合并逻辑
    ws['!merges'] = [
        { s: {r:0, c:0}, e: {r:0, c:17} }
    ];
    
    // 纵向合并前8列的基础属性表头
    for(let c=0; c<=7; c++) {
        ws['!merges'].push({ s: {r:3, c:c}, e: {r:4, c:c} });
    }
    
    // 横向合并各发电厂的表头标题
    let colIdx = 8;
    targetGencos.forEach(() => {
        ws['!merges'].push({ s: {r:3, c:colIdx}, e: {r:3, c:colIdx+1} });
        colIdx += 2;
    });
    
    // 6. 精细调整各列宽度
    ws['!cols'] = [
        {wpx: 40}, {wpx: 120}, {wpx: 120}, {wpx: 120}, {wpx: 80}, {wpx: 80}, {wpx: 80}, {wpx: 80}
    ];
    for(let i=0; i<10; i++) ws['!cols'].push({wpx: 90});
    
    XLSX.utils.book_append_sheet(wb, ws, '공동입찰 종합결과');
    
    // 触发下载
    const batchName = document.getElementById('batchName').value.trim() || '입찰평가';
    const date = new Date().toISOString().split('T')[0];
    const filename = `공동입찰_종합결과_${batchName}_${date}.xlsx`;
    
    XLSX.writeFile(wb, filename);
    showToast('Excel 파일이 성공적으로 내보내졌습니다!', 'success');
}

// ==================== 모달 제어 ====================

function closeModal(modalId) {
    document.getElementById(modalId).classList.remove('active');
}

document.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal')) {
        e.target.classList.remove('active');
    }
});

document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        document.querySelectorAll('.modal.active').forEach(modal => {
            modal.classList.remove('active');
        });
    }
    
    if (e.key === 'Enter' && !document.getElementById('loginOverlay').classList.contains('hidden')) {
        handleLogin();
    }
});
