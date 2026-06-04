// ProfitMatrix Application Logic

// State Management
const State = {
    customRates: {}, // { '2024-01': 5.30, ... }
    rawWorkbook: null, // Cache workbook to re-process on rate change
    processedData: {
        volumes: {},
        financials: {},
        categories: [],
        totalCompany: ''
    },
    ui: {
        activeTab: 'dashboard',
        startYear: '2026',
        startMonth: '01',
        endYear: '2026',
        endMonth: '12',
        selectedCategory: ''
    },
    charts: {
        trend: null,
        cat: null,
        unitTrend: null,
        rate: null,
        simTrend: null
    },
    unitAnalysis: {
        selectedCategory: '',
        selectedAccounts: ['매출액', '영업이익']
    },
    currentDashboardRev: [],
    currentUnitRev: [],
    currentUnitAmounts: {}
};

// Map Excel column indices to YYYY-MM
const colRatesMapping = {};
const offsets = [0, 2, 4, 8, 10, 12, 16, 18, 20, 24, 26, 28];
function buildColMapping() {
    let yearBases = [ {y: 2024, b: 1}, {y: 2025, b: 35}, {y: 2026, b: 69} ];
    yearBases.forEach(yb => {
        for(let i=0; i<12; i++) {
            let col = yb.b + offsets[i];
            let key = `${yb.y}-${String(i+1).padStart(2, '0')}`;
            colRatesMapping[col] = key;
        }
    });
}
buildColMapping();

// UI Elements
const els = {
    navItems: document.querySelectorAll('.nav-item'),
    sections: document.querySelectorAll('.view-section'),
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('excelFile'),
    logList: document.getElementById('logList'),
    uploadLog: document.getElementById('uploadLog'),
    statusBadge: document.getElementById('validationStatus'),
    statusText: document.querySelector('.status-text'),
    
    startYear: document.getElementById('startYear'),
    startMonth: document.getElementById('startMonth'),
    endYear: document.getElementById('endYear'),
    endMonth: document.getElementById('endMonth'),
    categoryFilter: document.getElementById('categoryFilter'),
    
    kpiRevenue: document.getElementById('kpiRevenue'),
    kpiVarCost: document.getElementById('kpiVarCost'),
    kpiFixCost: document.getElementById('kpiFixCost'),
    kpiOpProfit: document.getElementById('kpiOpProfit'),
    kpiVolume: document.getElementById('kpiVolume'),

    rateTbody: document.getElementById('rateTbody'),
    btnSaveRates: document.getElementById('btnSaveRates'),

    uaStartYear: document.getElementById('uaStartYear'),
    uaStartMonth: document.getElementById('uaStartMonth'),
    uaEndYear: document.getElementById('uaEndYear'),
    uaEndMonth: document.getElementById('uaEndMonth'),
    uaCategoryFilter: document.getElementById('uaCategoryFilter'),
    uaAccountGrid: document.getElementById('uaAccountGrid'),

    simRateSlider: document.getElementById('simRateSlider'),
    simRateVal: document.getElementById('simRateVal'),
    simOpProfit: document.getElementById('simOpProfit'),
    simOpProfitDiff: document.getElementById('simOpProfitDiff'),
    detectedIssuesGrid: document.getElementById('detectedIssuesGrid'),
    dangerCount: document.getElementById('dangerCount'),
    warningCount: document.getElementById('warningCount'),
    infoCount: document.getElementById('infoCount'),
    simTableBody: document.getElementById('simTableBody')
};

function init() {
    // Determine Rates
    const defaultRates = {
        '2022-01': 5.21, '2022-02': 5.23, '2022-03': 5.22, '2022-04': 5.25, '2022-05': 5.28, '2022-06': 5.30, '2022-07': 5.31, '2022-08': 5.33, '2022-09': 5.35, '2022-10': 5.37, '2022-11': 5.32, '2022-12': 5.38,
        '2023-01': 5.29, '2023-02': 5.31, '2023-03': 5.33, '2023-04': 5.34, '2023-05': 5.33, '2023-06': 5.36, '2023-07': 5.37, '2023-08': 5.40, '2023-09': 5.42, '2023-10': 5.45, '2023-11': 5.41, '2023-12': 5.39,
        '2024-01': 5.41, '2024-02': 5.44, '2024-03': 5.38, '2024-04': 5.44, '2024-05': 5.37, '2024-06': 5.42, '2024-07': 5.45, '2024-08': 5.41, '2024-09': 5.41, '2024-10': 5.43, '2024-11': 5.49, '2024-12': 5.64,
        '2025-01': 5.75, '2025-02': 5.69, '2025-03': 5.70, '2025-04': 5.58, '2025-05': 5.37, '2025-06': 5.24, '2025-07': 5.26, '2025-08': 5.29, '2025-09': 5.27, '2025-10': 5.40, '2025-11': 5.53, '2025-12': 5.57,
        '2026-01': 5.55, '2026-02': 5.57, '2026-03': 5.66
    };

    const savedRates = localStorage.getItem('profitMatrixRates_V2');
    if (savedRates && savedRates !== '{}') {
        try { State.customRates = JSON.parse(savedRates); } catch(e) { State.customRates = defaultRates; }
    } else {
        State.customRates = {...defaultRates};
        localStorage.setItem('profitMatrixRates_V2', JSON.stringify(State.customRates));
    }
    
    initSettingsTab();
    bindEvents();
    initCharts();
    
    // Sync state with DOM select values on load
    if (els.startYear) State.ui.startYear = els.startYear.value;
    if (els.startMonth) State.ui.startMonth = els.startMonth.value;
    if (els.endYear) State.ui.endYear = els.endYear.value;
    if (els.endMonth) State.ui.endMonth = els.endMonth.value;

    // Attempt cache load of processed data
    const savedData = localStorage.getItem('profitMatrixData_V5');
    if (savedData) {
        try {
            State.processedData = JSON.parse(savedData);
            processCategories();
            renderAccountCheckboxes();
            updateDashboard();
            updateUnitAnalysis();
            updateAnalysisReport();
            logMsg('Loaded previously saved parsed data.', 'log-info');
            setStatus(true, 'Data Loaded from Cache');
        } catch(e) { }
    }
    
    updateRateChart();
}

function initSettingsTab() {
    let html = '';
    for(let y=2022; y<=2026; y++) {
        html += `<tr><td>${y}년</td>`;
        for(let m=1; m<=12; m++) {
            let key = `${y}-${String(m).padStart(2, '0')}`;
            let val = State.customRates[key];
            let displayVal = val ? Number(val).toFixed(2) : '';
            html += `<td><input type="number" step="0.01" class="rate-input" id="rate_${key}" value="${displayVal}"></td>`;
        }
        html += `</tr>`;
    }
    els.rateTbody.innerHTML = html;
}

function bindEvents() {
    els.navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            switchTab(e.currentTarget.dataset.tab);
        });
    });
    
    let btnOpenReport = document.getElementById('btnOpenReport');
    if (btnOpenReport) {
        btnOpenReport.addEventListener('click', () => {
            switchTab('report');
        });
    }

    els.dropzone.addEventListener('dragover', (e) => { e.preventDefault(); els.dropzone.classList.add('dragover'); });
    els.dropzone.addEventListener('dragleave', (e) => { e.preventDefault(); els.dropzone.classList.remove('dragover'); });
    els.dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        els.dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFileUpload(e.dataTransfer.files[0]);
    });
    els.fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) handleFileUpload(e.target.files[0]);
    });

    els.btnSaveRates.addEventListener('click', () => {
        // Collect rates
        for(let y=2022; y<=2026; y++) {
            for(let m=1; m<=12; m++) {
                let key = `${y}-${String(m).padStart(2, '0')}`;
                let el = document.getElementById(`rate_${key}`);
                if (el && el.value.trim() !== '') {
                    State.customRates[key] = Number(el.value);
                } else {
                    delete State.customRates[key];
                }
            }
        }
        localStorage.setItem('profitMatrixRates_V2', JSON.stringify(State.customRates));
        alert('✓ 환율 설정이 성공적으로 저장되었습니다.\n(참고: 기존에 업로드된 파일을 다시 한 번 파싱합니다.)');
        
        updateRateChart();

        // Re-process if workbook is in memory
        if (State.rawWorkbook) {
            parseWorkbook(State.rawWorkbook);
        }
    });

    const handleDateChange = (source, event) => {
        let sy, sm, ey, em;
        if (source === 'dashboard') {
            sy = els.startYear.value; sm = els.startMonth.value;
            ey = els.endYear.value; em = els.endMonth.value;
        } else {
            sy = els.uaStartYear.value; sm = els.uaStartMonth.value;
            ey = els.uaEndYear.value; em = els.uaEndMonth.value;
        }
        
        let sIdx = (parseInt(sy) - 2024) * 12 + (parseInt(sm) - 1);
        let eIdx = (parseInt(ey) - 2024) * 12 + (parseInt(em) - 1);
        
        if (sIdx > eIdx) {
            if (event && event.target) {
                let tid = event.target.id.toLowerCase();
                if (tid.includes('end')) {
                    sy = ey; sm = em;
                } else {
                    ey = sy; em = sm;
                }
            } else {
                ey = sy; em = sm;
            }
        }
        
        State.ui.startYear = sy; State.ui.startMonth = sm;
        State.ui.endYear = ey; State.ui.endMonth = em;
        
        els.startYear.value = sy; els.startMonth.value = sm;
        els.endYear.value = ey; els.endMonth.value = em;
        els.uaStartYear.value = sy; els.uaStartMonth.value = sm;
        els.uaEndYear.value = ey; els.uaEndMonth.value = em;
        
        updateDashboard();
        updateUnitAnalysis();
        updateAnalysisReport();
    };

    els.startYear.addEventListener('change', (e) => handleDateChange('dashboard', e));
    els.startMonth.addEventListener('change', (e) => handleDateChange('dashboard', e));
    els.endYear.addEventListener('change', (e) => handleDateChange('dashboard', e));
    els.endMonth.addEventListener('change', (e) => handleDateChange('dashboard', e));
    els.categoryFilter.addEventListener('change', () => {
        let val = els.categoryFilter.value;
        State.ui.selectedCategory = val;
        els.uaCategoryFilter.value = val;
        State.unitAnalysis.selectedCategory = val;
        // Account reset removed to preserve user selection
        updateDashboard();
        renderAccountCheckboxes();
    });

    els.uaStartYear.addEventListener('change', (e) => handleDateChange('ua', e));
    els.uaStartMonth.addEventListener('change', (e) => handleDateChange('ua', e));
    els.uaEndYear.addEventListener('change', (e) => handleDateChange('ua', e));
    els.uaEndMonth.addEventListener('change', (e) => handleDateChange('ua', e));
    
    els.uaCategoryFilter.addEventListener('change', (e) => {
        let val = e.target.value;
        State.unitAnalysis.selectedCategory = val;
        // Account reset removed to preserve user selection
        els.categoryFilter.value = val;
        State.ui.selectedCategory = val;
        renderAccountCheckboxes();
        updateDashboard();
    });

    // Synchronize scroll between Chart and Table in Unit Analysis, and make Y-Axis sticky
    setTimeout(() => {
        const chartWrapper = document.querySelector('#unitAnalysisView .chart-scroll-wrapper');
        const tableWrapper = document.querySelector('#unitAnalysisView .table-scroll-wrapper');
        
        if (chartWrapper && tableWrapper) {
            let isSyncing = false;
            
            const syncScroll = (source, target) => {
                if (!isSyncing) {
                    isSyncing = true;
                    const sourceMaxScroll = source.scrollWidth - source.clientWidth;
                    const targetMaxScroll = target.scrollWidth - target.clientWidth;
                    if (sourceMaxScroll > 0 && targetMaxScroll > 0) {
                        const ratio = source.scrollLeft / sourceMaxScroll;
                        target.scrollLeft = ratio * targetMaxScroll;
                    } else {
                        target.scrollLeft = source.scrollLeft;
                    }
                    isSyncing = false;
                }
            };
            
            const makeElementSticky = (el) => {
                if (!el) return;
                let origX = el.getAttribute('data-orig-translate-x');
                if (origX === null) {
                    let transformAttr = el.getAttribute('transform') || '';
                    let match = transformAttr.match(/translate\(([-\d.]+)/);
                    origX = match ? parseFloat(match[1]) : 0;
                    el.setAttribute('data-orig-translate-x', origX);
                }
                origX = parseFloat(origX);
                el.setAttribute('transform', `translate(${origX + chartWrapper.scrollLeft}, 0)`);
            };
            
            const makeYAxisSticky = () => {
                const yAxis = chartWrapper.querySelector('.apexcharts-yaxis');
                if (yAxis) {
                    let bgRect = yAxis.querySelector('.sticky-yaxis-bg');
                    if (!bgRect) {
                        bgRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
                        bgRect.setAttribute('class', 'sticky-yaxis-bg');
                        bgRect.setAttribute('x', '-60');
                        bgRect.setAttribute('y', '-20');
                        bgRect.setAttribute('width', '80');
                        bgRect.setAttribute('height', '500');
                        bgRect.setAttribute('fill', '#0B0F19');
                        yAxis.insertBefore(bgRect, yAxis.firstChild);
                        
                        // Force Y-Axis group to render last (on top of chart lines)
                        yAxis.parentNode.appendChild(yAxis);
                        
                        // Also force Y-Axis title to render last if it exists as a sibling
                        const yAxisTitle = chartWrapper.querySelector('.apexcharts-yaxis-title');
                        if (yAxisTitle) {
                            yAxisTitle.parentNode.appendChild(yAxisTitle);
                        }
                    }
                    makeElementSticky(yAxis);
                    
                    const yAxisTitle = chartWrapper.querySelector('.apexcharts-yaxis-title');
                    if (yAxisTitle && !yAxis.contains(yAxisTitle)) {
                        makeElementSticky(yAxisTitle);
                    }
                }
            };
            
            chartWrapper.addEventListener('scroll', () => {
                syncScroll(chartWrapper, tableWrapper);
                makeYAxisSticky();
            });
            
            tableWrapper.addEventListener('scroll', () => {
                syncScroll(tableWrapper, chartWrapper);
                makeYAxisSticky();
            });
        }
    }, 200);

    if (els.simRateSlider) {
        els.simRateSlider.addEventListener('input', (e) => {
            let val = parseInt(e.target.value);
            runExchangeSimulation(val);
        });
    }
}

function switchTab(tabId) {
    els.navItems.forEach(i => i.classList.remove('active'));
    let navItem = document.querySelector(`[data-tab="${tabId}"]`);
    if (navItem) navItem.classList.add('active');
    
    els.sections.forEach(s => s.classList.remove('active'));
    document.getElementById(`${tabId}View`).classList.add('active');
    
    if (tabId === 'report') {
        updateAnalysisReport();
    }
    
    // Allow the DOM to update 'display: block' before resizing charts
    setTimeout(() => {
        window.dispatchEvent(new Event('resize'));
    }, 50);
}

function handleFileUpload(file) {
    els.uploadLog.classList.remove('hidden');
    els.logList.innerHTML = '';
    logMsg(`Processing ${file.name}...`, 'log-info');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            State.rawWorkbook = workbook;
            parseWorkbook(workbook);
        } catch (error) {
            logMsg(`Error reading Excel: ${error.message}`, 'log-error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseWorkbook(workbook) {
    const pData = { volumes: {}, financials: {}, categories: [], totalCompany: '' };

    if (!workbook.Sheets['매출&손익']) {
        logMsg("Critical Error: '매출&손익' sheet not found.", 'log-error');
        return;
    }

    // Determine target categories first
    const ignoreSheets = ['매출&손익', '재료비 등', '&'];
    workbook.SheetNames.forEach(sheetName => {
        let trimmedName = sheetName.trim();
        let isIgnored = ignoreSheets.includes(trimmedName) || 
                        trimmedName.includes('요약') || 
                        trimmedName.includes('#');
        if (!isIgnored) pData.categories.push(trimmedName);
    });

    if (pData.categories.length > 0) {
        pData.totalCompany = pData.categories.includes('성우비나()') ? '성우비나()' : pData.categories[0];
    } else {
        pData.totalCompany = 'Total';
    }
    
    const salesSheet = XLSX.utils.sheet_to_json(workbook.Sheets['매출&손익'], {header: 1, defval: 0});
    
    // Safely extract Volumes by searching arrays for Category names
    let readingVolumes = false;
    let readingVolumesHeaderStr = '';
    for (let i = 0; i < salesSheet.length; i++) {
        let row = salesSheet[i];
        if (!row) continue;
        
        // Find the Volume section header
        let rowStr = Object.values(row).join(' ').replace(/\s+/g, '');
        if (rowStr.includes('매출수량') || rowStr.includes('K/EA') || rowStr.includes('KEA')) {
            readingVolumes = true;
            readingVolumesHeaderStr = rowStr;
            continue;
        }
        
        // Stop if we reach the next metric section (e.g. Price/단가)
        let firstFewCols = String(row[0] || '') + String(row[1] || '') + String(row[2] || '');
        let cleanFirstCols = firstFewCols.replace(/\s+/g, '');
        
        const stopKeywords = ['단가', 'price', '매출액', '비용', 'cost', '영업이익', '매출원가', '재료비', '노무비', '판매관리비', '판관비', '환율', '인건비'];
        if (readingVolumes && !rowStr.includes('매출수량') && stopKeywords.some(kw => cleanFirstCols.toLowerCase().includes(kw))) {
            readingVolumes = false;
            continue;
        }
        
        if (readingVolumes) {
            let targetIdx = -1;
            let catStr = '';
            
            for (let checkCol = 0; checkCol <= 3; checkCol++) {
                let val = row[checkCol];
                if (val !== undefined && val !== null && val !== '') {
                    let valClean = String(val).replace(/\s+/g, '').toLowerCase();
                    
                    let matchedCat = pData.categories.find(c => {
                        let cClean = c.replace(/\s+/g, '').toLowerCase();
                        if (cClean === valClean) return true;
                        
                        let cCore = cClean.replace(/^mobile\(/, '').replace(/\)$/, '').replace(/^press\(/, '').replace(/\)$/, '');
                        let valCore = valClean.replace(/_kg|kg$/, '');
                        
                        let cNorm = cCore.replace(/[^a-z0-9가-힣]/g, '');
                        let vNorm = valCore.replace(/[^a-z0-9가-힣]/g, '');
                        
                        if (cNorm && vNorm && cNorm === vNorm) return true;
                        
                        if (vNorm.length >= 3 && cNorm.length >= 3) {
                            if (cNorm.includes(vNorm) || vNorm.includes(cNorm)) return true;
                        }
                        return false;
                    });
                    
                    if (matchedCat) {
                        targetIdx = checkCol;
                        catStr = matchedCat;
                        break;
                    }
                }
            }
            
            if (targetIdx !== -1) {
                let multiplier = 1; // 사용자가 수량이 실제 EA 단위(예: 67599993)라고 명시하였으므로 1로 고정
                
                // 매출&손익 layout is dense, relative to the category name index!
                let volArr = [];
                const getNum = (v) => { let n = Number(String(v).replace(/,/g,'').trim()); return isNaN(n)?0:n; };
                
                // 2024: 12 months starting right next to the label (targetIdx + 1..12)
                for(let c=1; c<=12; c++) volArr.push(getNum(row[targetIdx + c]) * multiplier);
                
                // 2025: Skips the 2024 '합계' column (targetIdx + 13). Starts at + 14..25
                for(let c=14; c<=25; c++) volArr.push(getNum(row[targetIdx + c]) * multiplier);
                
                // 2026: Skips the 2025 '합계' column (targetIdx + 26). Starts at + 27..38
                for(let c=27; c<=38; c++) volArr.push(getNum(row[targetIdx + c]) * multiplier);
                
                pData.volumes[catStr] = volArr; 
            }
        }
    }
    
    // Now process financials for each category sheet
    const majorCategories = ['매출액', '매출이익', '변동제조비', '변동판매비', '한계이익', '고정제조비', '고정판매비', '일반관리비', '손익분기점', '영업이익', '영업외손익'];

    pData.categories.forEach(trimmedName => {
        pData.financials[trimmedName] = {};

        let originalSheetName = workbook.SheetNames.find(s => s.trim() === trimmedName) || trimmedName;
        let sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[originalSheetName], {header: 1, defval: 0});
        
        let currentMajor = null;
        let currentSub = null;
        const subCategories = ['재료비', '노무비', '경비', '인건비'];

        for (let r = 4; r < sheetRows.length; r++) {
            let row = sheetRows[r];
            let accountRaw = row[0];
            if (typeof accountRaw === 'string' && accountRaw.trim() !== '') {
                let accName = accountRaw.replace(/\s+/g, ''); 
                
                // Track hierarchy context
                if (majorCategories.includes(accName)) {
                    currentMajor = accName;
                    currentSub = null;
                } else if (subCategories.includes(accName)) {
                    currentSub = accName;
                }
                
                // Make unique key based on context to prevent overwriting (e.g. 변동제조비_경비 vs 고정제조비_경비)
                let uniqueKey = accName;
                let level = 1; // 1: major, 2: sub, 3: item
                
                if (!majorCategories.includes(accName)) {
                    if (subCategories.includes(accName)) {
                        uniqueKey = currentMajor + '_' + accName;
                        level = 2;
                    } else {
                        uniqueKey = currentMajor + '_' + (currentSub ? currentSub + '_' : '') + accName;
                        level = 3;
                    }
                }
                
                // Prevent duplicate keys overwriting each other
                let baseKey = uniqueKey;
                let counter = 1;
                while (pData.financials[trimmedName][uniqueKey]) {
                    uniqueKey = baseKey + '_' + counter;
                    counter++;
                }

                let convertedRow = row.map((val, colIdx) => {
                    if (colIdx === 0) return val;
                    
                    let rate = 5.3;
                    let mKey = colRatesMapping[colIdx];
                    if (mKey && State.customRates[mKey] > 0) {
                        rate = State.customRates[mKey];
                    }

                    let numVal = Number(String(val).replace(/,/g,'').trim());
                    // 원화 환산: Raw VND * (rate / 100) = Raw KRW
                    if (!isNaN(numVal)) return numVal * rate / 100;
                    return 0;
                });
                
                pData.financials[trimmedName][uniqueKey] = {
                    label: accountRaw.replace(/\s+/g, ''),
                    key: uniqueKey,
                    major: currentMajor,
                    level: level,
                    data: convertedRow
                };
            }
        }
    });

    State.processedData = pData;
    validateData(pData);

    localStorage.setItem('profitMatrixData_V5', JSON.stringify(pData));
    processCategories();
    updateAnalysisReport();
    switchTab('dashboard');
    alert("✓ 엑셀 파싱 및 커스텀 환율 변환 완료!\n대시보드가 정상적으로 업데이트 되었습니다.");
}

function validateData(pData) {
    let tc = pData.totalCompany;
    if (!pData.financials[tc]) {
        setStatus(false, "Total Sheet Missing");
        return;
    }
    let isValid = true;
    let sum = 0;
    pData.categories.forEach(cat => {
        if (cat !== tc) {
            if (pData.financials[cat]['매출액']) sum += Number(pData.financials[cat]['매출액'].data[1]) || 0;
        }
    });
    let totalVal = pData.financials[tc] && pData.financials[tc]['매출액'] ? Number(pData.financials[tc]['매출액'].data[1]) || 0 : 0;
    let diffPercent = totalVal !== 0 ? Math.abs((totalVal - sum) / totalVal) : 0;
    
    if (diffPercent > 0.01) setStatus(false, "Validation Failed");
    else setStatus(true, "Data Verified & Synced");
}

function getCustomChartColors(seriesArray) {
    let customColors = [];
    let defaultColors = ['#06b6d4', '#ec4899', '#6366f1', '#f59e0b', '#8b5cf6', '#a855f7', '#3b82f6'];
    seriesArray.forEach((series, idx) => {
        let name = series.name.replace(/\s+/g, '');
        if (name.includes('매출액') || name.includes('매출단가')) {
            customColors.push('#ef4444'); // Bright Red
        } else if (name.includes('영업이익')) {
            customColors.push('#10b981'); // Bright Emerald Green
        } else {
            customColors.push(defaultColors[idx % defaultColors.length]);
        }
    });
    return customColors;
}

function formatCategoryLabel(rawCat, forChart = false) {
    let displayCat = rawCat === 'Mobile(SC)' ? 'Mobile(S/C)' : rawCat;
    displayCat = displayCat.replace(/\(\)$/, ''); // Remove empty trailing parenthesis
    if (forChart) {
        let shortLabelMatch = displayCat.match(/\((.*?)\)/);
        if (shortLabelMatch && shortLabelMatch[1].trim() !== '') {
            return shortLabelMatch[1];
        }
    }
    return displayCat;
}

function processCategories() {
    let html = '';
    let uaHtml = '';
    let tc = State.processedData.totalCompany;
    
    if (State.processedData.categories.includes(tc)) {
        let displayName = tc === '성우비나()' ? '베트남법인' : tc;
        html += `<option value="${tc}" selected>${displayName}</option>`;
        uaHtml += `<option value="${tc}" selected>${displayName}</option>`;
    }
    State.processedData.categories.forEach(cat => {
        if(cat !== tc) {
            let displayCat = formatCategoryLabel(cat, false);
            html += `<option value="${cat}">${displayCat}</option>`;
            uaHtml += `<option value="${cat}">${displayCat}</option>`;
        }
    });
    els.categoryFilter.innerHTML = html;
    els.uaCategoryFilter.innerHTML = uaHtml;
    
    let defaultCat = State.processedData.categories.includes(tc) ? tc : State.processedData.categories[0];
    if (defaultCat) {
        State.ui.selectedCategory = defaultCat;
        els.categoryFilter.value = defaultCat;
        State.unitAnalysis.selectedCategory = defaultCat;
        els.uaCategoryFilter.value = defaultCat;
        State.unitAnalysis.selectedAccounts = ['매출액', '영업이익'];
    }
    
    renderAccountCheckboxes();
}

function extractPeriodicData(rowArr) {
    if (!rowArr) return new Array(36).fill(0);
    let allData = [];
    const getNum = (val) => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        let p = Number(String(val).replace(/,/g, '').trim());
        return isNaN(p) ? 0 : p;
    };
    // Extract using the known offsets exactly for 2024, 2025, 2026 chunks
    for(let i=0; i<12; i++) { allData.push(getNum(rowArr[1 + offsets[i]])); }
    for(let i=0; i<12; i++) { allData.push(getNum(rowArr[35 + offsets[i]])); }
    for(let i=0; i<12; i++) { allData.push(getNum(rowArr[69 + offsets[i]])); }
    return allData;
}

function updateDashboard() {
    if (!State.processedData.categories.length) return;

    let cat = State.ui.selectedCategory;
    
    let sIdx = (parseInt(State.ui.startYear) - 2024) * 12 + (parseInt(State.ui.startMonth) - 1);
    let eIdx = (parseInt(State.ui.endYear) - 2024) * 12 + (parseInt(State.ui.endMonth) - 1);
    
    sIdx = Math.max(0, Math.min(35, sIdx));
    eIdx = Math.max(0, Math.min(35, eIdx));
    if (sIdx > eIdx) { let temp = sIdx; sIdx = eIdx; eIdx = temp; }
    
    let labels = [];
    for(let i=sIdx; i<=eIdx; i++) {
        let y = 2024 + Math.floor(i / 12);
        let m = (i % 12) + 1;
        labels.push(`${y}년 ${m}월`);
    }

    let revArr = [], varCostArr = [], fixCostArr = [], opProfitArr = [], volArr = [];

    if (cat === 'All') cat = State.processedData.totalCompany;

    let sgaArr = [];
    if (State.processedData.financials[cat]) {
        let fin = State.processedData.financials[cat];
        revArr = extractPeriodicData(fin['매출액']?.data).slice(sIdx, eIdx + 1);
        varCostArr = extractPeriodicData(fin['변동제조비']?.data).slice(sIdx, eIdx + 1);
        fixCostArr = extractPeriodicData(fin['고정제조비']?.data).slice(sIdx, eIdx + 1);
        opProfitArr = extractPeriodicData(fin['영업이익']?.data).slice(sIdx, eIdx + 1);
        
        let varSArr = extractPeriodicData(fin['변동판매비']?.data).slice(sIdx, eIdx + 1);
        let fixSArr = extractPeriodicData(fin['고정판매비']?.data).slice(sIdx, eIdx + 1);
        let admArr = extractPeriodicData(fin['일반관리비']?.data).slice(sIdx, eIdx + 1);
        sgaArr = new Array(eIdx - sIdx + 1).fill(0);
        for(let i=0; i<sgaArr.length; i++) sgaArr[i] = (varSArr[i]||0) + (fixSArr[i]||0) + (admArr[i]||0);
    }
    
    if (State.processedData.volumes[cat] && State.processedData.volumes[cat].reduce((a,b)=>a+b, 0) > 0) {
        volArr = State.processedData.volumes[cat].slice(sIdx, eIdx + 1);
    } else if (cat === State.processedData.totalCompany) {
        volArr = new Array(eIdx - sIdx + 1).fill(0);
        Object.keys(State.processedData.volumes).forEach(k => {
            if (k !== State.processedData.totalCompany) {
                let vArr = State.processedData.volumes[k].slice(sIdx, eIdx + 1);
                for(let i=0; i<volArr.length; i++) volArr[i] += vArr[i] || 0;
            }
        });
    } else {
        volArr = new Array(eIdx - sIdx + 1).fill(0);
    }

    State.currentDashboardRev = revArr;

    let totalRev = revArr.reduce((a,b)=>a+b, 0);
    let totalVarCost = varCostArr.reduce((a,b)=>a+b, 0);
    let totalFixCost = fixCostArr.reduce((a,b)=>a+b, 0);
    let totalSGA = sgaArr.reduce((a,b)=>a+b, 0);
    let totalOpProfit = opProfitArr.reduce((a,b)=>a+b, 0);
    let totalVol = volArr.reduce((a,b)=>a+b, 0);

    let getRatioHtml = (val) => `<span style="font-size: 0.95rem; color: var(--text-secondary); font-weight: normal; margin-left: 4px;">(${totalRev ? (val / totalRev * 100).toFixed(1) : '0.0'}%)</span>`;

    if (els.kpiRevenue) els.kpiRevenue.innerHTML = formatCurr(totalRev);
    if (els.kpiVarCost) els.kpiVarCost.innerHTML = `${formatCurr(totalVarCost)}${getRatioHtml(totalVarCost)}`;
    if (els.kpiFixCost) els.kpiFixCost.innerHTML = `${formatCurr(totalFixCost)}${getRatioHtml(totalFixCost)}`;
    if (document.getElementById('kpiSGA')) document.getElementById('kpiSGA').innerHTML = `${formatCurr(totalSGA)}${getRatioHtml(totalSGA)}`;
    if (els.kpiOpProfit) els.kpiOpProfit.innerHTML = `${formatCurr(totalOpProfit)}${getRatioHtml(totalOpProfit)}`;
    if (els.kpiVolume) els.kpiVolume.innerText = formatNum(totalVol) + ' EA';

    let yearAggr = {};
    labels.forEach((lbl, idx) => {
        let yr = lbl.split('년')[0];
        if (!yearAggr[yr]) yearAggr[yr] = { rev: 0, varCost: 0, fixCost: 0, sga: 0, opProfit: 0, vol: 0 };
        yearAggr[yr].rev += revArr[idx];
        yearAggr[yr].varCost += varCostArr[idx];
        yearAggr[yr].fixCost += fixCostArr[idx];
        yearAggr[yr].sga += sgaArr[idx];
        yearAggr[yr].opProfit += opProfitArr[idx];
        yearAggr[yr].vol += volArr[idx];
    });

    let bdRevHTML = '', bdVarCostHTML = '', bdFixCostHTML = '', bdSgaHTML = '', bdOpProfHTML = '', bdVolHTML = '';
    Object.keys(yearAggr).forEach(yr => {
        let yrRev = yearAggr[yr].rev;
        let yrRatio = (val) => ` <small style="color: var(--text-secondary);">(${yrRev ? (val / yrRev * 100).toFixed(1) : '0.0'}%)</small>`;
        bdRevHTML += `<div class="yr-row"><span>${yr}년:</span> <span>${formatCurr(yrRev)}</span></div>`;
        bdVarCostHTML += `<div class="yr-row"><span>${yr}년:</span> <span>${formatCurr(yearAggr[yr].varCost)}${yrRatio(yearAggr[yr].varCost)}</span></div>`;
        bdFixCostHTML += `<div class="yr-row"><span>${yr}년:</span> <span>${formatCurr(yearAggr[yr].fixCost)}${yrRatio(yearAggr[yr].fixCost)}</span></div>`;
        bdSgaHTML += `<div class="yr-row"><span>${yr}년:</span> <span>${formatCurr(yearAggr[yr].sga)}${yrRatio(yearAggr[yr].sga)}</span></div>`;
        bdOpProfHTML += `<div class="yr-row"><span>${yr}년:</span> <span>${formatCurr(yearAggr[yr].opProfit)}${yrRatio(yearAggr[yr].opProfit)}</span></div>`;
        bdVolHTML += `<div class="yr-row"><span>${yr}년:</span> <span>${formatNum(yearAggr[yr].vol)} EA</span></div>`;
    });

    if (document.getElementById('bdRevenue')) document.getElementById('bdRevenue').innerHTML = bdRevHTML;
    if (document.getElementById('bdVarCost')) document.getElementById('bdVarCost').innerHTML = bdVarCostHTML;
    if (document.getElementById('bdFixCost')) document.getElementById('bdFixCost').innerHTML = bdFixCostHTML;
    if (document.getElementById('bdSGA')) document.getElementById('bdSGA').innerHTML = bdSgaHTML;
    if (document.getElementById('bdOpProfit')) document.getElementById('bdOpProfit').innerHTML = bdOpProfHTML;
    if (document.getElementById('bdVolume')) document.getElementById('bdVolume').innerHTML = bdVolHTML;

    let trendSeries = [];
    if (State.unitAnalysis.selectedAccounts && State.unitAnalysis.selectedAccounts.length > 0) {
        State.unitAnalysis.selectedAccounts.forEach(acc => {
            if (State.processedData.financials[cat]) {
                let accObj = State.processedData.financials[cat][acc];
                if (accObj && accObj.data) {
                    let extracted = extractPeriodicData(accObj.data).slice(sIdx, eIdx + 1);
                    trendSeries.push({ name: accObj.label.replace(/\s+/g, '') + ' (원)', data: extracted });
                }
            }
        });
    }

    let trendCanvas = document.querySelector("#trendChart");
    if (trendCanvas) {
        trendCanvas.style.width = '100%';
    }

    // Generate beautifully grouped year categories centered under month periods
    let rawLabels = [];
    for(let i=sIdx; i<=eIdx; i++) {
        let y = 2024 + Math.floor(i / 12);
        let m = (i % 12) + 1;
        rawLabels.push({ year: y, month: `${m}월` });
    }
    let yearGroups = {};
    rawLabels.forEach((item, idx) => {
        if (!yearGroups[item.year]) yearGroups[item.year] = [];
        yearGroups[item.year].push(idx);
    });
    let chartLabels = [];
    rawLabels.forEach((item, idx) => {
        let group = yearGroups[item.year];
        let midIdx = group[Math.floor((group.length - 1) / 2)];
        if (idx === midIdx) {
            chartLabels.push([item.month, `${item.year}년`]);
        } else {
            chartLabels.push(item.month);
        }
    });

    // Add "평균" (Average) label
    chartLabels.push(['평균', '선택기간']);

    // Calculate and append average for each series
    let periodLength = eIdx - sIdx + 1;
    if (periodLength > 0) {
        let validMonthsCount = 0;
        for (let i = 0; i < periodLength; i++) {
            if (State.currentDashboardRev[i] > 0) validMonthsCount++;
        }
        let div = validMonthsCount > 0 ? validMonthsCount : periodLength;

        trendSeries.forEach(series => {
            let sum = 0;
            for (let i = 0; i < periodLength; i++) {
                if (State.currentDashboardRev[i] > 0 || validMonthsCount === 0) {
                    sum += series.data[i];
                }
            }
            series.data.push(sum / div);
        });
        
        // Also add the average to State.currentDashboardRev so tooltips work correctly
        let totalRevSum = 0;
        for (let i = 0; i < periodLength; i++) {
            if (State.currentDashboardRev[i] > 0 || validMonthsCount === 0) {
                totalRevSum += State.currentDashboardRev[i];
            }
        }
        State.currentDashboardRev.push(totalRevSum / div);
    }

    State.charts.trend.updateOptions({
        colors: getCustomChartColors(trendSeries),
        xaxis: { type: 'category', categories: chartLabels, tickPlacement: 'on' },
        series: trendSeries
    });

    let rawCatNames = [];
    let displayCatLabels = [];
    let catRevs = [];
    State.processedData.categories.forEach(c => {
        if (c !== State.processedData.totalCompany) {
            rawCatNames.push(c);
            displayCatLabels.push(formatCategoryLabel(c, true));
            
            let cRev = 0;
            if (State.processedData.financials[c] && State.processedData.financials[c]['매출액']) {
                let rArr = extractPeriodicData(State.processedData.financials[c]['매출액'].data).slice(sIdx, eIdx + 1);
                cRev = rArr.reduce((a,b)=>a+b, 0);
            }
            catRevs.push(cRev);
        }
    });
    State.currentCatRev = catRevs;

    let newCatSeries = [];
    if (State.unitAnalysis.selectedAccounts && State.unitAnalysis.selectedAccounts.length > 0) {
        State.unitAnalysis.selectedAccounts.forEach(acc => {
            let label = acc;
            if (State.processedData.financials[State.processedData.totalCompany] && State.processedData.financials[State.processedData.totalCompany][acc]) {
                label = State.processedData.financials[State.processedData.totalCompany][acc].label;
            }
            
            let dataArr = [];
            rawCatNames.forEach(c => {
                let val = 0;
                if (State.processedData.financials[c] && State.processedData.financials[c][acc]) {
                    let arr = extractPeriodicData(State.processedData.financials[c][acc].data).slice(sIdx, eIdx + 1);
                    val = arr.reduce((a,b)=>a+b, 0);
                }
                dataArr.push(val);
            });
            newCatSeries.push({ name: label + ' (원)', data: dataArr });
        });
    }

    if (State.charts.cat) {
        State.charts.cat.updateOptions({
            colors: getCustomChartColors(newCatSeries),
            xaxis: { type: 'category', categories: displayCatLabels, tickPlacement: 'on' },
            series: newCatSeries
        });
    }
}

function renderAccountCheckboxes() {
    let cat = State.unitAnalysis.selectedCategory;
    els.uaAccountGrid.innerHTML = '';
    if (!cat || !State.processedData.financials[cat]) return;

    let accounts = Object.values(State.processedData.financials[cat]);
    let html = '';
    
    // Group accounts by Major to form a tree structure
    let majorGroups = {};
    accounts.forEach(acc => {
        if (!acc.major || acc.major === '영업외손익') return;
        
        // Force fix alignment levels for cached data
        let cleanLabel = acc.label.replace(/\s+/g, '');
        const subCategories = ['재료비', '노무비', '경비', '인건비'];
        const majorCategories = ['매출액', '매출이익', '변동제조비', '변동판매비', '한계이익', '고정제조비', '고정판매비', '일반관리비', '손익분기점', '영업이익', '영업외손익'];
        if (majorCategories.includes(cleanLabel)) {
            acc.level = 1;
        } else if (subCategories.includes(cleanLabel)) {
            acc.level = 2;
        } else {
            acc.level = 3;
        }

        if (!majorGroups[acc.major]) majorGroups[acc.major] = [];
        majorGroups[acc.major].push(acc);
    });

    Object.keys(majorGroups).forEach(majorKey => {
        let groupItems = majorGroups[majorKey];
        // Render Major Parent
        let majorItemsHtml = '';
        
        groupItems.forEach(item => {
            if (item.label.includes('합계') || item.label.length < 2) return;
            
            let isChecked = State.unitAnalysis.selectedAccounts.includes(item.key) ? 'checked' : '';
            let indentClass = item.level === 1 ? 'level-1' : (item.level === 2 ? 'level-2' : 'level-3');
            let isParent = (item.level === 1 && groupItems.length > 1) || (item.level === 2 && groupItems.filter(i => i.key.startsWith(item.key)).length > 1);
            let parentAttr = isParent ? `data-parent="${item.key}"` : '';
            let immediateParentKey = item.level > 1 ? item.key.substring(0, item.key.lastIndexOf('_')) : '';
            
            let toggleBtn = isParent ? `<span class="toggle-btn" data-toggle="${item.key}">+</span>` : `<span class="toggle-btn empty"></span>`;
            let displayStyle = item.level > 1 ? 'style="display: none;"' : '';

            majorItemsHtml += `
                <div class="account-item-wrapper ${indentClass}" data-key="${item.key}" data-parent-key="${immediateParentKey}" ${displayStyle}>
                    ${toggleBtn}
                    <label class="account-tree-item" style="margin: 0; flex: 1;">
                        <input type="checkbox" value="${item.key}" ${isChecked} ${parentAttr}>
                        <span class="label-text">${item.label.replace(/\s+/g, '')}</span>
                    </label>
                </div>
            `;
        });
        
        html += `<div class="tree-group">${majorItemsHtml}</div>`;
    });

    els.uaAccountGrid.innerHTML = html;

    // Accordion Logic
    els.uaAccountGrid.querySelectorAll('.toggle-btn:not(.empty)').forEach(btn => {
        btn.addEventListener('click', (e) => {
            let targetKey = e.target.getAttribute('data-toggle');
            let isExpanded = e.target.classList.contains('expanded');
            
            if (isExpanded) {
                e.target.classList.remove('expanded');
                e.target.innerText = '+';
                els.uaAccountGrid.querySelectorAll('.account-item-wrapper').forEach(el => {
                    let k = el.getAttribute('data-key');
                    if (k !== targetKey && k.startsWith(targetKey + '_')) {
                        el.style.display = 'none';
                        let childBtn = el.querySelector('.toggle-btn');
                        if (childBtn && !childBtn.classList.contains('empty')) {
                            childBtn.classList.remove('expanded');
                            childBtn.innerText = '+';
                        }
                    }
                });
            } else {
                e.target.classList.add('expanded');
                e.target.innerText = '-';
                els.uaAccountGrid.querySelectorAll(`.account-item-wrapper[data-parent-key="${targetKey}"]`).forEach(el => {
                    el.style.display = 'flex';
                });
            }
        });
    });

    // Checkbox Tree Logic
    els.uaAccountGrid.querySelectorAll('input[type="checkbox"]').forEach(chk => {
        chk.addEventListener('change', (e) => {
            let val = e.target.value;
            let group = e.target.closest('.tree-group');
            let isChecking = e.target.checked;
            
            // Auto-check logic removed as per user request (parents don't auto-select children)

            // Sync State with DOM checks
            let newlySelected = [];
            els.uaAccountGrid.querySelectorAll('input[type="checkbox"]:checked').forEach(c => {
                newlySelected.push(c.value);
            });
            
            State.unitAnalysis.selectedAccounts = newlySelected;

            if (State.unitAnalysis.selectedAccounts.length === 0) {
                State.unitAnalysis.selectedAccounts = ['매출액', '영업이익'];
                renderAccountCheckboxes();
                return;
            }
            updateCheckboxIndeterminateStates();
            updateUnitAnalysis();
            updateDashboard();
        });
    });

    function updateCheckboxIndeterminateStates() {
        let allCheckboxes = els.uaAccountGrid.querySelectorAll('input[type="checkbox"]');
        allCheckboxes.forEach(chk => {
            let key = chk.value;
            let hasSelectedChild = Array.from(allCheckboxes).some(c => c.checked && c.value !== key && c.value.startsWith(key + '_'));
            chk.indeterminate = hasSelectedChild && !chk.checked;
        });
    }

    updateCheckboxIndeterminateStates();
    updateUnitAnalysis();
    updateDashboard();
}

function updateUnitAnalysis() {
    let cat = State.unitAnalysis.selectedCategory;
    if (!cat || !State.processedData.financials[cat]) return;

    let sIdx = (parseInt(State.ui.startYear) - 2024) * 12 + (parseInt(State.ui.startMonth) - 1);
    let eIdx = (parseInt(State.ui.endYear) - 2024) * 12 + (parseInt(State.ui.endMonth) - 1);
    
    sIdx = Math.max(0, Math.min(35, sIdx));
    eIdx = Math.max(0, Math.min(35, eIdx));

    let labels = [];
    for(let i=sIdx; i<=eIdx; i++) {
        let y = 2024 + Math.floor(i / 12);
        let m = (i % 12) + 1;
        labels.push(`${y}년 ${m}월`);
    }

    let periodLength = eIdx - sIdx + 1;
    let volArr = new Array(periodLength).fill(0);
    if (State.processedData.volumes[cat] && State.processedData.volumes[cat].reduce((a,b)=>a+b, 0) > 0) {
        volArr = State.processedData.volumes[cat].slice(sIdx, eIdx + 1);
    } else if (cat === State.processedData.totalCompany) {
        volArr = new Array(periodLength).fill(0);
        Object.keys(State.processedData.volumes).forEach(k => {
            if (k !== State.processedData.totalCompany) {
                let vArr = State.processedData.volumes[k].slice(sIdx, eIdx + 1);
                for(let i=0; i<volArr.length; i++) volArr[i] += vArr[i] || 0;
            }
        });
    } else {
        volArr = new Array(periodLength).fill(0);
    }

    let revArr = new Array(periodLength).fill(0);
    if (State.processedData.financials[cat] && State.processedData.financials[cat]['매출액']) {
        revArr = extractPeriodicData(State.processedData.financials[cat]['매출액'].data).slice(sIdx, eIdx + 1);
    }
    
    let unitRevArr = new Array(periodLength).fill(0);
    for(let i=0; i<periodLength; i++) {
        if (volArr[i] > 0) unitRevArr[i] = revArr[i] / volArr[i];
    }
    State.currentUnitRev = unitRevArr;
    State.currentUnitAmounts = {};

    let unitSeries = [];
    
    State.unitAnalysis.selectedAccounts.forEach(acc => {
        let accObj = State.processedData.financials[cat][acc];
        if (accObj && accObj.data) {
            let extracted = extractPeriodicData(accObj.data).slice(sIdx, eIdx + 1);
            State.currentUnitAmounts[accObj.label] = extracted;
            
            let unitArr = new Array(periodLength).fill(0);
            for(let i=0; i<periodLength; i++) {
                if (volArr[i] && volArr[i] > 0) {
                    unitArr[i] = extracted[i] / volArr[i];
                } else {
                    unitArr[i] = 0;
                }
            }
            let cleanLabel = accObj.label.replace(/\s+/g, '');
            let seriesLabel = cleanLabel === '매출액' ? '매출단가' : cleanLabel;
            unitSeries.push({ name: seriesLabel, data: unitArr });
        }
    });

    const chartWrapper = document.querySelector('#unitAnalysisView .chart-scroll-wrapper');
    const tableWrapper = document.querySelector('#unitAnalysisView .table-scroll-wrapper');
    if (chartWrapper) chartWrapper.scrollLeft = 0;
    if (tableWrapper) tableWrapper.scrollLeft = 0;

    let unitCanvas = document.querySelector("#unitTrendChart");

    // Generate beautifully grouped year categories centered under month periods
    let rawLabels = [];
    for(let i=sIdx; i<=eIdx; i++) {
        let y = 2024 + Math.floor(i / 12);
        let m = (i % 12) + 1;
        rawLabels.push({ year: y, month: `${m}월` });
    }
    let yearGroups = {};
    rawLabels.forEach((item, idx) => {
        if (!yearGroups[item.year]) yearGroups[item.year] = [];
        yearGroups[item.year].push(idx);
    });
    let chartLabels = [];
    rawLabels.forEach((item, idx) => {
        let group = yearGroups[item.year];
        let midIdx = group[Math.floor((group.length - 1) / 2)];
        if (idx === midIdx) {
            chartLabels.push([item.month, `${item.year}년`]);
        } else {
            chartLabels.push(item.month);
        }
    });

    // Add "평균" (Average) label
    chartLabels.push(['평균', '선택기간']);
    
    // Calculate and append average for each series
    if (periodLength > 0) {
        let validMonthsCount = 0;
        for (let i = 0; i < periodLength; i++) {
            if (State.currentUnitRev[i] > 0) validMonthsCount++;
        }
        let div = validMonthsCount > 0 ? validMonthsCount : periodLength;

        unitSeries.forEach(series => {
            let sum = 0;
            for (let i = 0; i < periodLength; i++) {
                if (State.currentUnitRev[i] > 0 || validMonthsCount === 0) {
                    sum += series.data[i];
                }
            }
            series.data.push(sum / div);
        });
        
        let totalUnitRevSum = 0;
        for (let i = 0; i < periodLength; i++) {
            if (State.currentUnitRev[i] > 0 || validMonthsCount === 0) {
                totalUnitRevSum += State.currentUnitRev[i];
            }
        }
        State.currentUnitRev.push(totalUnitRevSum / div);
    }

    if (State.charts.unitTrend) {
        State.charts.unitTrend.updateOptions({
            colors: getCustomChartColors(unitSeries),
            xaxis: { type: 'category', categories: chartLabels, tickPlacement: 'on' },
            series: unitSeries
        });
    }

    // Update Unit Table
    let tableThead = document.querySelector('#unitTable thead');
    let tableBody = document.getElementById('unitTableBody');
    if (tableThead && tableBody) {
        let years = {};
        labels.forEach(lbl => {
            let parts = lbl.split(' ');
            let yStr = parts[0];
            let mStr = parts[1];
            if(!years[yStr]) years[yStr] = [];
            years[yStr].push(mStr);
        });

        let head1 = '<tr><th rowspan="3" style="min-width: 120px; vertical-align: middle;">구분 (원/EA)</th>';
        let head2 = '<tr>';
        let head3 = '<tr>';
        
        Object.keys(years).forEach(yStr => {
            let months = years[yStr];
            head1 += `<th colspan="${months.length * 2}">${yStr}</th>`;
            months.forEach(mStr => {
                head2 += `<th colspan="2">${mStr}</th>`;
                head3 += `<th style="min-width: 40px; font-size: 0.8rem; padding: 4px 2px;">단가</th><th style="min-width: 35px; font-size: 0.8rem; padding: 4px 2px;">비율</th>`;
            });
        });
        
        // Add "Average" column group
        head1 += `<th colspan="2" style="color: var(--accent-cyan);">선택기간</th>`;
        head2 += `<th colspan="2" style="color: var(--accent-cyan);">평균</th>`;
        head3 += `<th style="min-width: 40px; font-size: 0.8rem; padding: 4px 2px; color: var(--accent-cyan);">단가</th><th style="min-width: 35px; font-size: 0.8rem; padding: 4px 2px; color: var(--accent-cyan);">비율</th>`;

        head1 += '</tr>';
        head2 += '</tr>';
        head3 += '</tr>';
        
        tableThead.innerHTML = head1 + head2 + head3;

        let bodyHtml = '';
        unitSeries.forEach(series => {
            let displayName = series.name === '매출액' ? '매출단가' : series.name;
            bodyHtml += `<tr><td style="text-align: left; padding-left: 16px;">${displayName}</td>`;
            series.data.forEach((val, idx) => {
                let isAvg = (idx === series.data.length - 1);
                let rev = State.currentUnitRev[idx];
                let ratioStr = '';
                if (series.name !== '매출액' && rev && rev > 0) {
                    ratioStr = ((val / rev) * 100).toFixed(1) + '%';
                }
                let cellColor = isAvg ? 'color: var(--accent-cyan); font-weight: 500;' : '';
                let ratioColor = isAvg ? 'color: rgba(6, 182, 212, 0.8);' : 'color: var(--text-secondary);';
                bodyHtml += `<td style="text-align: right; padding-right: 4px; min-width: 40px; font-size: 0.85rem; ${cellColor}">${formatUnit(val)}</td>`;
                bodyHtml += `<td style="text-align: right; padding-right: 4px; min-width: 35px; font-size: 0.8rem; ${ratioColor}">${ratioStr}</td>`;
            });
            bodyHtml += `</tr>`;
        });
        tableBody.innerHTML = bodyHtml;
    }

    // Make chart dynamic so it scrolls proportionately when the period is large
    let canvasEl = document.querySelector("#unitTrendChart");
    if (canvasEl) {
        let chartMinWidth = Math.max((periodLength + 1) * 80, 800); // Give 80px per month, ensuring a minimum of 800px
        canvasEl.style.minWidth = chartMinWidth + 'px';
        canvasEl.style.width = '100%';
    }
}

function initCharts() {
    const opts = { 
        chart: { foreColor: '#94a3b8', toolbar: { show: false }, zoom: { enabled: false }, parentHeightOffset: 0 }, 
        theme: { mode: 'dark' }, 
        grid: { show: true, borderColor: 'rgba(255, 255, 255, 0.5)', strokeDashArray: 2, yaxis: { lines: { show: true } }, padding: { top: 15, bottom: 0, left: 30, right: 30 } }, 
        legend: { show: true, position: 'bottom', height: 40 } 
    };
    
    let ttTrend = { theme: 'dark', y: { formatter: function(val, { dataPointIndex }) {
        let rev = State.currentDashboardRev[dataPointIndex];
        let ratio = (rev && rev > 0) ? ((val / rev) * 100).toFixed(1) + '%' : '0%';
        if (val === rev) return `금액: ₩ ${Math.round(val).toLocaleString()}`;
        return `금액: ₩ ${Math.round(val).toLocaleString()} (비율: ${ratio})`;
    }}};

    State.charts.trend = new ApexCharts(document.querySelector("#trendChart"), {
        ...opts,
        tooltip: ttTrend,
        series: [],
        chart: { ...opts.chart, type: 'line', height: 350 },
        colors: ['#06b6d4', '#ec4899', '#6366f1', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444'],
        markers: { size: 5, hover: { size: 7 } },
        dataLabels: { enabled: false },
        stroke: { curve: 'straight', width: 3 },
        xaxis: {
            type: 'category',
            tickPlacement: 'on',
            categories: [],
            labels: {
                rotate: 0,
                rotateAlways: false,
                style: { colors: '#94a3b8', fontSize: '11px' }
            }
        },
        yaxis: { title: { text: '금액 (원)' }, labels: { formatter: (val) => val === 0 ? '0' : Math.round(val / 100000000).toLocaleString() + '억' } }
    });
    State.charts.trend.render();

    let ttCat = { theme: 'dark', y: { formatter: function(val, { dataPointIndex, seriesIndex, w }) {
        let rev = State.currentCatRev ? State.currentCatRev[dataPointIndex] : 0;
        let ratio = (rev && rev > 0) ? ((val / rev) * 100).toFixed(1) + '%' : '0.0%';
        let sName = w.globals.seriesNames[seriesIndex] || '';
        
        if (sName.includes('매출액') || sName.includes('매출단가')) return `금액: ₩ ${Math.round(val).toLocaleString()}`;
        return `금액: ₩ ${Math.round(val).toLocaleString()} (비율: ${ratio})`;
    }}};

    State.charts.cat = new ApexCharts(document.querySelector("#catChart"), {
        ...opts, tooltip: ttCat, series: [], chart: { ...opts.chart, type: 'line', height: 350 }, colors: ['#06b6d4', '#ec4899', '#6366f1', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444'], markers: { size: 5, hover: { size: 7 } }, stroke: { curve: 'straight', width: 3 }, dataLabels: { enabled: false }, xaxis: { type: 'category', tickPlacement: 'on', categories: [], labels: { rotate: -45, hideOverlappingLabels: false, trim: true, style: { fontSize: '10px', colors: '#94a3b8' } } }, yaxis: { title: { text: '금액 (원)' }, labels: { formatter: (val) => val === 0 ? '0' : Math.round(val / 100000000).toLocaleString() + '억' } }
    });
    State.charts.cat.render();

    let ttUnit = { theme: 'dark', y: { formatter: function(val, { dataPointIndex, seriesIndex, w }) {
        let rev = State.currentUnitRev[dataPointIndex];
        let ratio = (rev && rev > 0) ? ((val / rev) * 100).toFixed(1) + '%' : '0.0%';
        let sName = w.globals.seriesNames[seriesIndex];
        
        if (sName === '매출액') return `단가: ₩ ${formatUnit(val)}`;
        return `단가: ₩ ${formatUnit(val)} (비율: ${ratio})`;
    }}};

    State.charts.unitTrend = new ApexCharts(document.querySelector("#unitTrendChart"), {
        ...opts,
        grid: {
            ...opts.grid,
            padding: {
                ...opts.grid.padding,
                left: 45 // Increased left padding to prevent first label from being cut off
            }
        },
        tooltip: ttUnit,
        series: [],
        chart: { ...opts.chart, type: 'line', height: 300 },
        colors: ['#06b6d4', '#ec4899', '#6366f1', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444'],
        markers: { size: 5, hover: { size: 7 } },
        dataLabels: { enabled: false },
        stroke: { curve: 'straight', width: 3 },
        xaxis: {
            type: 'category',
            tickPlacement: 'on',
            categories: [],
            labels: {
                rotate: 0,
                rotateAlways: false,
                style: { colors: '#94a3b8', fontSize: '11px' }
            }
        },
        yaxis: { title: { text: '단가 (원)' }, labels: { formatter: (val) => Math.round(val).toLocaleString() } }
    });
    State.charts.unitTrend.render();

    State.charts.rate = new ApexCharts(document.querySelector("#rateChart"), {
        ...opts, tooltip: { theme: 'dark' }, series: [{ name: '환율 (원/동*100)', data: [] }], chart: { ...opts.chart, type: 'area', height: 250 }, colors: ['#f59e0b'], fill: { type: 'gradient', gradient: { shadeIntensity: 1, opacityFrom: 0.4, opacityTo: 0.05, stops: [0, 100] } }, markers: { size: 3, hover: { size: 5 } }, dataLabels: { enabled: false }, stroke: { curve: 'smooth', width: 2 }, xaxis: { type: 'category', tickPlacement: 'on', categories: [], tickAmount: 12, labels: { style: { colors: '#94a3b8' } } }, yaxis: { min: 5.00, max: 6.00, title: { text: '환율 (KRW/VND*100)' }, labels: { formatter: (val) => val.toFixed(2) } }
    });
    State.charts.rate.render();

    State.charts.simTrend = new ApexCharts(document.querySelector("#simChart"), {
        ...opts,
        tooltip: { theme: 'dark', y: { formatter: (val) => '₩ ' + Math.round(val).toLocaleString('ko-KR') } },
        series: [{ name: '영업이익 (원)', data: [] }],
        chart: { ...opts.chart, type: 'bar', height: 200 },
        colors: ['#06b6d4'],
        plotOptions: { bar: { borderRadius: 4, columnWidth: '40%', distributed: true } },
        dataLabels: { enabled: false },
        xaxis: { type: 'category', categories: ['현재 환율', '가상 환율'] },
        yaxis: { labels: { formatter: (val) => Math.round(val).toLocaleString() } }
    });
    State.charts.simTrend.render();
}

// ==========================================
// Analysis Report & Issue Detection Engine
// ==========================================
function updateAnalysisReport() {
    if (!els.detectedIssuesGrid) return;
    
    let reportPeriodEl = document.getElementById('reportPeriodText');
    if (reportPeriodEl) {
        reportPeriodEl.innerText = `분석 대상 기간: ${State.ui.startYear}년 ${State.ui.startMonth}월 ~ ${State.ui.endYear}년 ${State.ui.endMonth}월`;
    }

    if (!State.processedData.categories.length) {
        els.dangerCount.innerText = '0';
        els.warningCount.innerText = '0';
        els.infoCount.innerText = '0';
        els.detectedIssuesGrid.innerHTML = `
            <div class="no-data-placeholder glass-panel" style="grid-column: 1 / -1; padding: 40px; text-align: center;">
                <span style="font-size: 2.5rem; display: block; margin-bottom: 16px;">📂</span>
                <h3>분석할 수 있는 엑셀 데이터가 없습니다.</h3>
                <p style="color: var(--text-secondary); margin-top: 8px;">데이터 업로드 탭에서 ERP 로우 데이터를 먼저 업로드해 주세요.</p>
            </div>
        `;
        return;
    }

    let { issues, dangerCount, warningCount, infoCount } = runIssueDetection();

    // Update Counters
    els.dangerCount.innerText = dangerCount;
    els.warningCount.innerText = warningCount;
    els.infoCount.innerText = infoCount;

    // Render Issue Cards
    if (issues.length === 0) {
        els.detectedIssuesGrid.innerHTML = `
            <div class="no-data-placeholder glass-panel" style="grid-column: 1 / -1; padding: 40px; text-align: center; border-color: rgba(16, 185, 129, 0.3);">
                <span style="font-size: 2.5rem; display: block; margin-bottom: 16px;">✅</span>
                <h3>감지된 중대한 리스크가 없습니다.</h3>
                <p style="color: var(--text-secondary); margin-top: 8px;">현재 설정된 필터 조건에서 모든 부문이 양호한 재무 상태를 유지하고 있습니다.</p>
            </div>
        `;
    } else {
        let html = '';
        issues.forEach(issue => {
            html += `
                <div class="issue-card glass-panel" style="border-left: 4px solid var(--accent-${issue.level === 'danger' ? 'red' : (issue.level === 'warning' ? 'magenta' : 'cyan')});">
                    <div class="issue-card-header">
                        <span class="issue-title" style="font-weight:600; font-size:1.05rem;">${issue.title}</span>
                        <span class="issue-badge ${issue.level}">${issue.badge}</span>
                    </div>
                    <p class="issue-desc" style="margin-top: 8px; color: var(--text-secondary); line-height: 1.5; font-size:0.9rem;">${issue.desc}</p>
                    <div class="issue-action-box" style="margin-top:12px;">
                        <strong style="color: var(--text-primary);">🛠️ ${issue.action.title}</strong>
                        <p style="margin-top:4px; color: var(--text-secondary); font-size:0.85rem; line-height: 1.4;">${issue.action.p}</p>
                    </div>
                </div>
            `;
        });
        els.detectedIssuesGrid.innerHTML = html;
    }

    // Trigger Simulator with current slider value
    let currentSliderVal = els.simRateSlider ? parseInt(els.simRateSlider.value) : 0;
    runExchangeSimulation(currentSliderVal);

    // Resize simulator chart since it might have been hidden
    setTimeout(() => {
        if (State.charts.simTrend) {
            window.dispatchEvent(new Event('resize'));
        }
    }, 100);
}

function runIssueDetection() {
    let issues = [];
    let dangerCount = 0;
    let warningCount = 0;
    let infoCount = 0;

    let sIdx = (parseInt(State.ui.startYear) - 2024) * 12 + (parseInt(State.ui.startMonth) - 1);
    let eIdx = (parseInt(State.ui.endYear) - 2024) * 12 + (parseInt(State.ui.endMonth) - 1);
    sIdx = Math.max(0, Math.min(35, sIdx));
    eIdx = Math.max(0, Math.min(35, eIdx));
    let periodLength = eIdx - sIdx + 1;

    let tc = State.processedData.totalCompany;

    State.processedData.categories.forEach(cat => {
        if (cat === tc) return;

        let fin = State.processedData.financials[cat];
        if (!fin) return;

        let sales = extractPeriodicData(fin['매출액']?.data).slice(sIdx, eIdx + 1);
        let varMfg = extractPeriodicData(fin['변동제조비']?.data).slice(sIdx, eIdx + 1);
        
        let varSell = new Array(periodLength).fill(0);
        Object.keys(fin).forEach(k => {
            if (k.startsWith('변동판매비')) {
                let d = extractPeriodicData(fin[k]?.data).slice(sIdx, eIdx + 1);
                for(let i=0; i<periodLength; i++) varSell[i] += d[i] || 0;
            }
        });

        let fixedMfg = extractPeriodicData(fin['고정제조비']?.data).slice(sIdx, eIdx + 1);
        
        let fixedSell = new Array(periodLength).fill(0);
        let genAdmin = new Array(periodLength).fill(0);
        Object.keys(fin).forEach(k => {
            if (k.startsWith('고정판매비')) {
                let d = extractPeriodicData(fin[k]?.data).slice(sIdx, eIdx + 1);
                for(let i=0; i<periodLength; i++) fixedSell[i] += d[i] || 0;
            }
            if (k.startsWith('일반관리비')) {
                let d = extractPeriodicData(fin[k]?.data).slice(sIdx, eIdx + 1);
                for(let i=0; i<periodLength; i++) genAdmin[i] += d[i] || 0;
            }
        });

        let totalSales = sales.reduce((a, b) => a + b, 0);
        let totalVarMfg = varMfg.reduce((a, b) => a + b, 0);
        let totalVarSell = varSell.reduce((a, b) => a + b, 0);
        let totalFixedMfg = fixedMfg.reduce((a, b) => a + b, 0);
        let totalFixedSell = fixedSell.reduce((a, b) => a + b, 0);
        let totalGenAdmin = genAdmin.reduce((a, b) => a + b, 0);

        let totalVariable = totalVarMfg + totalVarSell;
        let totalFixed = totalFixedMfg + totalFixedSell + totalGenAdmin;
        let marginalProfit = totalSales - totalVariable;
        let marginalProfitRatio = totalSales ? (marginalProfit / totalSales * 100).toFixed(1) : '0.0';

        // 1. Check if Marginal Loss exists
        if (totalSales > 0 && marginalProfit < 0) {
            dangerCount++;
            issues.push({
                level: 'danger',
                badge: '위험',
                title: `품목분류 [${cat}] 한계적자 발생`,
                desc: `조회 기간 동안 총 매출액은 ${formatCurr(totalSales)}이나, 변동비가 매출을 초과하여 ${formatCurr(totalVariable)} (${((totalVariable/totalSales)*100).toFixed(1)}%)에 달해 한계이익 적자 ${formatCurr(marginalProfit)} (한계이익률 ${marginalProfitRatio}%)를 기록하고 있습니다. 이는 생산을 늘릴수록 손실이 누적되는 적자수주 상태입니다.`,
                action: {
                    title: '추천 조치 계획',
                    p: '• 1단계: 부자재 및 물류비 소요 단위 BOM 정밀 재검증 및 단가 긴급 재협상<br>• 2단계: 자동화 공정 재배치를 통한 현장 생산인력 효율 개선 및 직접 변동비 축소<br>• 3단계: 단가 보전이 안 되는 적자 제품의 수주량을 통제하고 고수익 품목으로 전환 생산'
                }
            });
        }

        // 2. Check if Fixed Cost Burden is too high (> 30%)
        let fixedRatio = totalSales ? (totalFixed / totalSales) : 0;
        if (totalSales > 0 && fixedRatio > 0.30 && marginalProfit >= 0) {
            warningCount++;
            issues.push({
                level: 'warning',
                badge: '주의',
                title: `품목분류 [${cat}] 고정비 부담 과다 경보`,
                desc: `조회 기간 동안 매출액 ${formatCurr(totalSales)} 대비 고정비가 ${formatCurr(totalFixed)} (${(fixedRatio*100).toFixed(1)}%)로 내부 경보 기준치(30%)를 초과하고 있습니다. 가동률 하락 시 바로 대규모 영업 손실로 이어질 우려가 큽니다.`,
                action: {
                    title: '추천 조치 계획',
                    p: '• 1단계: 유휴 설비 자산의 타 라인 이설 또는 처분을 통한 감가상각 부담 축소<br>• 2단계: 간접인력 최적 배치 및 고정 제경비 항목에 대한 제로베이스 원가 절감 정밀 실사<br>• 3단계: 생산 변동성에 유연하게 대처할 수 있도록 외주 가공 비중 전환 검토'
                }
            });
        }
    });

    // 3. Check for Tet holiday impact (2월)
    if (State.processedData.financials[tc]) {
        let fin = State.processedData.financials[tc];
        let salesData = extractPeriodicData(fin['매출액']?.data);

        let startY = parseInt(State.ui.startYear);
        let endY = parseInt(State.ui.endYear);

        for (let y = Math.max(2024, startY); y <= Math.min(2026, endY); y++) {
            let yBase = (y - 2024) * 12;
            let janSales = salesData[yBase + 0] || 0;
            let febSales = salesData[yBase + 1] || 0;
            let marSales = salesData[yBase + 2] || 0;

            let avgJanMar = (janSales + marSales) / 2;
            if (avgJanMar > 0 && febSales / avgJanMar < 0.75) {
                let dropPct = ((1 - febSales / avgJanMar) * 100).toFixed(1);
                warningCount++;
                issues.push({
                    level: 'warning',
                    badge: '주의',
                    title: `${y}년 2월 구정(Tet) 조업도 손실 분석`,
                    desc: `${y}년 2월 베트남 법인 매출액은 ${formatCurr(febSales)}로, 전후 월(1월, 3월) 평균 매출 ${formatCurr(avgJanMar)} 대비 ${dropPct}% 급감하였습니다. 베트남 구정 연휴(Tet) 동안 가동 중단에 따른 생산 비조업 손실이 주 원인입니다.`,
                    action: {
                        title: '추천 조치 계획',
                        p: '• 1단계: 구정 전후 집중 생산 추진 및 조기 재고 구축으로 출하 차질 최소화<br>• 2단계: 연휴 직후 조기 복귀 인센티브 지원 등을 통해 라인 셋업 시간을 단축하여 가동률 복구<br>• 3단계: 2월 한 달에 집중 배부되는 감가상각 고정비를 연간 조업도 배부 방식으로 보완 검토'
                    }
                });
            }
        }
    }

    // 4. Exchange Rate Sensitivity Analysis
    if (State.processedData.financials[tc]) {
        let fin = State.processedData.financials[tc];
        let sales = extractPeriodicData(fin['매출액']?.data).slice(sIdx, eIdx + 1);
        let opProfit = extractPeriodicData(fin['영업이익']?.data).slice(sIdx, eIdx + 1);

        let totalSales = sales.reduce((a, b) => a + b, 0);
        let totalOpProfit = opProfit.reduce((a, b) => a + b, 0);

        if (totalSales > 0) {
            infoCount++;
            let rateDeclineImpact = totalOpProfit * 0.1;

            issues.push({
                level: 'info',
                badge: '정보',
                title: '환율(KRW/USD) 변동에 따른 영업이익 민감도',
                desc: `현재 조회 기간의 베트남 법인 총 매출은 ${formatCurr(totalSales)}이며 원화 영업이익은 ${formatCurr(totalOpProfit)}입니다. 베트남 법인의 제품 판가(수출) 및 주요 원부자재 결제 통화는 미국 달러(USD) 기반이므로, KRW/USD 환율이 10% 상승할 시 원화 환산 영업이익이 약 ${formatCurr(rateDeclineImpact)} 증가하며, 10% 하락할 시 동일 금액만큼 감소하는 고민감 리스크 구조입니다.`,
                action: {
                    title: '추천 위험 관리 계획',
                    p: '• 1단계: KRW/USD 환율 가상 시뮬레이터를 활용한 원화 변동성 정기 모니터링 체계 가동<br>• 2단계: 현지 내수 매출 확대 및 결제 통화 포트폴리오 다변화로 특정 외환 의존도 감소<br>• 3단계: 환변동 보험 가입 또는 은행 연계 통화선도 계약을 통해 연간 경영계획 목표 손익 선제적 방어'
                }
            });
        }
    }

    // 5. Account-level Analysis for Selected Category and Period
    let selCat = State.ui.selectedCategory;
    if (selCat && State.processedData.financials[selCat]) {
        let fin = State.processedData.financials[selCat];
        let salesData = extractPeriodicData(fin['매출액']?.data).slice(sIdx, eIdx + 1);
        let totalSales = salesData.reduce((a, b) => a + b, 0);

        let totalCostsAll = 0;
        let accountStats = [];

        Object.keys(fin).forEach(accKey => {
            let accObj = fin[accKey];
            let isRevenueOrProfit = accObj.label.includes('매출') || accObj.label.includes('이익') || accObj.label.includes('손익');
            if (isRevenueOrProfit) return;

            let d = extractPeriodicData(accObj.data).slice(sIdx, eIdx + 1);
            let accTotal = d.reduce((a, b) => a + b, 0);
            
            if (accTotal <= 0) return;

            // Only consider lowest level accounts
            if (accObj.level === 3 || (accObj.level === 2 && !Object.keys(fin).some(k => fin[k].level === 3 && k.startsWith(accObj.key + '_')))) {
                totalCostsAll += accTotal;
                accountStats.push({
                    key: accKey,
                    label: accObj.label.replace(/\s+/g, ''),
                    data: d,
                    total: accTotal
                });
            }
        });

        accountStats.forEach(acc => {
            // Check 2: High Impact Account (>= 20% of total costs)
            let costRatio = totalCostsAll > 0 ? (acc.total / totalCostsAll) : 0;
            if (costRatio >= 0.20) {
                infoCount++;
                issues.push({
                    level: 'info',
                    badge: '정보',
                    title: `[${selCat}] 비중 과다 핵심 비용: ${acc.label}`,
                    desc: `해당 조회 기간 동안 '${acc.label}' 계정의 총 지출액은 ${formatCurr(acc.total)}으로, 전체 비용의 ${(costRatio * 100).toFixed(1)}%를 차지하는 핵심 비용입니다.`,
                    action: {
                        title: '추천 원가 관리 계획',
                        p: `• 1단계: ${acc.label} 발생 원인 및 단가/수량(P/Q) 변동 요인 정밀 분석<br>• 2단계: 동종 업계 벤치마킹 및 대체 공정/공급사 발굴 검토`
                    }
                });
            }

            // Check 1: Spike Detection (>= 100% increase over average)
            if (periodLength >= 2) {
                let currentMonthVal = acc.data[periodLength - 1];
                let previousMonthsSum = acc.total - currentMonthVal;
                let previousMonthsAvg = previousMonthsSum / (periodLength - 1);

                if (previousMonthsAvg > 0 && currentMonthVal >= previousMonthsAvg * 2.0) {
                    let spikeRatio = ((currentMonthVal - previousMonthsAvg) / previousMonthsAvg * 100).toFixed(1);
                    warningCount++;
                    issues.push({
                        level: 'warning',
                        badge: '주의',
                        title: `[${selCat}] 비용 급증 계정 경보: ${acc.label}`,
                        desc: `조회 기간 마지막 달의 '${acc.label}' 지출액이 ${formatCurr(currentMonthVal)}으로, 이전 기간 평균(${formatCurr(previousMonthsAvg)}) 대비 ${spikeRatio}% 급증하였습니다.`,
                        action: {
                            title: '추천 조치 계획',
                            p: `• 1단계: 전표 및 증빙 자료 대조를 통한 이상 지출 내역 점검<br>• 2단계: 일시적 비용 여부 파악 및 익월 비용 발생 통제 강화`
                        }
                    });
                }
            }

            // Check 3: Revenue Drop vs Cost Rise
            if (periodLength >= 2 && totalSales > 0) {
                let currentSales = salesData[periodLength - 1];
                let previousSales = salesData[periodLength - 2];
                let currentCost = acc.data[periodLength - 1];
                let previousCost = acc.data[periodLength - 2];

                if (previousSales > currentSales && currentCost > previousCost && previousCost > 0) {
                    dangerCount++;
                    issues.push({
                        level: 'danger',
                        badge: '위험',
                        title: `[${selCat}] 매출 둔화 대비 원가 상승 역전: ${acc.label}`,
                        desc: `전월 대비 매출은 감소(${formatCurr(previousSales)} → ${formatCurr(currentSales)})했음에도 불구하고, '${acc.label}' 계정의 지출액은 오히려 증가(${formatCurr(previousCost)} → ${formatCurr(currentCost)})하는 역전 현상이 감지되었습니다.`,
                        action: {
                            title: '긴급 원가 통제 계획',
                            p: `• 1단계: 조업도 하락에도 감소하지 않는 원인 파악 및 즉각적인 비용 집행 동결<br>• 2단계: 변동비 성격의 계정일 경우, 매출 연동형 예산 통제 시스템 도입`
                        }
                    });
                }
            }
        });
    }

    return { issues, dangerCount, warningCount, infoCount };
}

function runExchangeSimulation(percentChange) {
    let tc = State.processedData.totalCompany;
    if (!tc || !State.processedData.financials[tc]) return;

    let sIdx = (parseInt(State.ui.startYear) - 2024) * 12 + (parseInt(State.ui.startMonth) - 1);
    let eIdx = (parseInt(State.ui.endYear) - 2024) * 12 + (parseInt(State.ui.endMonth) - 1);
    sIdx = Math.max(0, Math.min(35, sIdx));
    eIdx = Math.max(0, Math.min(35, eIdx));

    let fin = State.processedData.financials[tc];
    let sales = extractPeriodicData(fin['매출액']?.data).slice(sIdx, eIdx + 1);
    let opProfit = extractPeriodicData(fin['영업이익']?.data).slice(sIdx, eIdx + 1);

    let originalSales = sales.reduce((a, b) => a + b, 0);
    let originalOpProfit = opProfit.reduce((a, b) => a + b, 0);

    let factor = 1 + (percentChange / 100);
    let simulatedSales = originalSales * factor;
    let simulatedOpProfit = originalOpProfit * factor;
    let profitDiff = simulatedOpProfit - originalOpProfit;

    // Update slider label
    let rateValText = `${percentChange > 0 ? '+' : ''}${percentChange}%`;
    if (percentChange === 0) {
        rateValText += ' (기본환율)';
        els.simRateVal.style.color = 'var(--text-secondary)';
    } else {
        rateValText += percentChange > 0 ? ' (USD 환율 상승 / 원화 약세 / 원화 환산 증가)' : ' (USD 환율 하락 / 원화 강세 / 원화 환산 감소)';
        els.simRateVal.style.color = percentChange > 0 ? 'var(--accent-green)' : 'var(--accent-red)';
    }
    els.simRateVal.innerText = rateValText;

    // Update simulated values
    els.simOpProfit.innerText = formatCurr(simulatedOpProfit);
    
    let diffPct = originalOpProfit ? (profitDiff / Math.abs(originalOpProfit) * 100).toFixed(1) : '0.0';
    let diffText = `${profitDiff >= 0 ? '+' : ''}${formatCurr(profitDiff)} (${profitDiff >= 0 ? '+' : ''}${diffPct}% 변동)`;
    els.simOpProfitDiff.innerText = diffText;
    els.simOpProfitDiff.style.color = profitDiff >= 0 ? 'var(--accent-green)' : 'var(--accent-red)';

    // Update ApexChart
    if (State.charts.simTrend) {
        State.charts.simTrend.updateSeries([{
            name: '영업이익 (원)',
            data: [
                { x: '현재 환율 (0%)', y: Math.round(originalOpProfit) },
                { x: `가상 환율 (${percentChange > 0 ? '+' : ''}${percentChange}%)`, y: Math.round(simulatedOpProfit) }
            ]
        }]);
    }

    // Update Scenario Table
    let tableHtml = '';
    const scenarios = [-10, -5, 0, 5, 10];
    
    let rowSales = `<tr><td style="text-align: left; font-weight: 500; padding-left:16px;">매출액</td>`;
    let rowProfit = `<tr><td style="text-align: left; font-weight: 500; padding-left:16px;">영업이익</td>`;
    let rowMargin = `<tr><td style="text-align: left; font-weight: 500; padding-left:16px;">영업이익률</td>`;
    let rowVariance = `<tr><td style="text-align: left; font-weight: 500; padding-left:16px;">영업이익 변동액</td>`;

    scenarios.forEach(sc => {
        let f = 1 + (sc / 100);
        let scSales = originalSales * f;
        let scProfit = originalOpProfit * f;
        let scMargin = scSales ? (scProfit / scSales * 100).toFixed(1) + '%' : '0.0%';
        let scVariance = scProfit - originalOpProfit;

        let activeStyle = sc === percentChange ? 'style="background: rgba(6, 182, 212, 0.15); font-weight: bold; color: var(--accent-cyan);"' : '';

        rowSales += `<td ${activeStyle}>${formatShort(scSales)}</td>`;
        rowProfit += `<td ${activeStyle}>${formatShort(scProfit)}</td>`;
        rowMargin += `<td ${activeStyle}>${scMargin}</td>`;
        
        let varText = sc === 0 ? '-' : `${scVariance >= 0 ? '+' : ''}${formatShort(scVariance)}`;
        let varColor = sc === 0 ? 'var(--text-secondary)' : (scVariance >= 0 ? 'var(--accent-green)' : 'var(--accent-red)');
        let tdStyle = sc === percentChange ? `style="background: rgba(6, 182, 212, 0.15); font-weight: bold; color: ${varColor};"` : `style="color: ${varColor}; font-family:monospace;"`;
        rowVariance += `<td ${tdStyle}>${varText}</td>`;
    });

    rowSales += '</tr>';
    rowProfit += '</tr>';
    rowMargin += '</tr>';
    rowVariance += '</tr>';

    els.simTableBody.innerHTML = rowSales + rowProfit + rowMargin + rowVariance;
}

function updateRateChart() {
    if (!State.charts.rate) return;
    let categories = [];
    let data = [];
    for(let y=2022; y<=2026; y++) {
        for(let m=1; m<=12; m++) {
            let key = `${y}-${String(m).padStart(2, '0')}`;
            categories.push(key);
            let val = State.customRates[key] || null;
            data.push(val);
        }
    }
    State.charts.rate.updateOptions({
        xaxis: { categories: categories },
        series: [{ name: '환율 (원/동*100)', data: data }]
    });
}

function logMsg(msg, cClass) {
    const li = document.createElement('li'); li.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`; li.className = cClass; els.logList.prepend(li);
}
function setStatus(isOk, text) {
    els.statusBadge.classList.remove('hidden');
    isOk ? els.statusBadge.classList.remove('error') : els.statusBadge.classList.add('error');
    els.statusText.innerText = text;
}
function formatCurr(val) { return '₩ ' + Math.round(val).toLocaleString('ko-KR'); }
function formatNum(val) { return Math.round(val).toLocaleString('ko-KR'); }
function formatUnit(val) { return Number(val).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); }
function formatShort(val) {
    if (val >= 100000000) return (val / 100000000).toFixed(1) + '억';
    if (val >= 1000000) return (val / 1000000).toFixed(1) + '백만';
    return Math.round(val).toLocaleString();
}
document.addEventListener('DOMContentLoaded', init);
