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
    compAccountGrid: document.getElementById('compAccountGrid'),

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
    initComparativeAnalysis();
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
    if (els.compAccountGrid) els.compAccountGrid.innerHTML = html;

    function bindGridEvents(grid) {
        if (!grid) return;
        
        // Accordion Logic
        grid.querySelectorAll('.toggle-btn:not(.empty)').forEach(btn => {
            btn.addEventListener('click', (e) => {
                let targetKey = e.target.getAttribute('data-toggle');
                let isExpanded = e.target.classList.contains('expanded');
                
                if (isExpanded) {
                    e.target.classList.remove('expanded');
                    e.target.innerText = '+';
                    grid.querySelectorAll('.account-item-wrapper').forEach(el => {
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
                    grid.querySelectorAll(`.account-item-wrapper[data-parent-key="${targetKey}"]`).forEach(el => {
                        el.style.display = 'flex';
                    });
                }
            });
        });

        // Checkbox Tree Logic
        grid.querySelectorAll('input[type="checkbox"]').forEach(chk => {
            chk.addEventListener('change', (e) => {
                // Read from THIS grid only
                let newlySelected = [];
                grid.querySelectorAll('input[type="checkbox"]:checked').forEach(c => {
                    newlySelected.push(c.value);
                });
                
                if (newlySelected.length === 0) {
                    State.unitAnalysis.selectedAccounts = ['매출액', '영업이익'];
                    renderAccountCheckboxes();
                    if (typeof runComparativeAnalysis === 'function') {
                        runComparativeAnalysis();
                    }
                    return;
                }
                
                State.unitAnalysis.selectedAccounts = newlySelected;
                
                // Sync ALL other instances directly
                [els.uaAccountGrid, els.compAccountGrid, document.getElementById('accountSelector')].forEach(g => {
                    if (g && g !== grid) {
                        g.querySelectorAll('input[type="checkbox"]').forEach(c => {
                            c.checked = newlySelected.includes(c.value);
                        });
                    }
                });
                
                updateCheckboxIndeterminateStates();
                updateUnitAnalysis();
                updateDashboard();
                if (typeof runComparativeAnalysis === 'function') {
                    runComparativeAnalysis();
                }
            });
        });
    }

    bindGridEvents(els.uaAccountGrid);
    bindGridEvents(els.compAccountGrid);

    function updateCheckboxIndeterminateStates() {
        [els.uaAccountGrid, els.compAccountGrid].forEach(grid => {
            if (!grid) return;
            let allCheckboxes = grid.querySelectorAll('input[type="checkbox"]');
            allCheckboxes.forEach(chk => {
                let key = chk.value;
                let hasSelectedChild = Array.from(allCheckboxes).some(c => c.checked && c.value !== key && c.value.startsWith(key + '_'));
                chk.indeterminate = hasSelectedChild && !chk.checked;
            });
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
    let selCat = State.ui.selectedCategory || tc;
    
    // --- 환율 변동성 민감도 고정 섹션 (항상 하단 표시) ---
    let exchangeRateSection = document.getElementById('exchangeRateSection');
    let exchangeRateSensitivityCard = document.getElementById('exchangeRateSensitivityCard');
    
    if (State.processedData.financials[tc]) {
        let fin = State.processedData.financials[tc];
        let sales = extractPeriodicData(fin['매출액']?.data).slice(sIdx, eIdx + 1);
        let opProfit = extractPeriodicData(fin['영업이익']?.data).slice(sIdx, eIdx + 1);
        let totalSales = sales.reduce((a, b) => a + b, 0);
        let totalOpProfit = opProfit.reduce((a, b) => a + b, 0);

        if (totalSales > 0 && exchangeRateSection && exchangeRateSensitivityCard) {
            exchangeRateSection.style.display = 'block';
            let rateDeclineImpact = totalOpProfit * 0.1;
            let sensitivityHtml = `
                <div class="issue-card glass-panel info">
                    <div class="issue-card-header">
                        <h4 class="issue-title"><span class="issue-badge info">정보</span>환율(KRW/USD) 변동에 따른 영업이익 민감도</h4>
                    </div>
                    <p class="issue-desc">현재 조회 기간의 베트남 법인 총 매출은 ${formatCurr(totalSales)}이며 원화 영업이익은 ${formatCurr(totalOpProfit)}입니다. 베트남 법인의 제품 판가(수출) 및 주요 원부자재 결제 통화는 미국 달러(USD) 기반이므로, KRW/USD 환율이 10% 상승할 시 원화 환산 영업이익이 약 ${formatCurr(rateDeclineImpact)} 증가하며, 10% 하락할 시 동일 금액만큼 감소하는 고민감 리스크 구조입니다.</p>
                    <div class="issue-action-box">
                        <strong>추천 위험 관리 계획</strong>
                        <p>• 1단계: 하단의 KRW/USD 환율 가상 시뮬레이터를 활용한 원화 변동성 정기 모니터링 체계 가동<br>• 2단계: 현지 내수 매출 확대 및 결제 통화 포트폴리오 다변화로 특정 외환 의존도 감소<br>• 3단계: 환변동 보험 가입 또는 은행 연계 통화선도 계약을 통해 연간 경영계획 목표 손익 선제적 방어</p>
                    </div>
                </div>
            `;
            exchangeRateSensitivityCard.innerHTML = sensitivityHtml;
        } else if (exchangeRateSection) {
            exchangeRateSection.style.display = 'none';
        }
    } else if (exchangeRateSection) {
        exchangeRateSection.style.display = 'none';
    }

    // --- 11대 핵심 지표 분석 엔진 (Option B: 문제 발생 시 카드 표시) ---
    if (selCat && State.processedData.financials[selCat]) {
        let fin = State.processedData.financials[selCat];
        
        let m = {
            sales: 0, valLoss: 0, grossProfit: 0, varMfg: 0, fixedMfg: 0, 
            sga: 0, material: 0, labor: 0, others: 0, marginalProfit: 0, 
            deprecMaintInd: 0, bep: 0, opProfit: 0,
            salesData: [], opProfitData: []
        };

        // 데이터 집계 (Aggregation)
        Object.keys(fin).forEach(accKey => {
            let accObj = fin[accKey];
            let lbl = accObj.label.replace(/\s+/g, '');
            let d = extractPeriodicData(accObj.data).slice(sIdx, eIdx + 1);
            let total = d.reduce((a, b) => a + b, 0);

            if (lbl === '매출액') { m.sales += total; m.salesData = d; }
            else if (lbl === '제품평가손실') m.valLoss += total;
            else if (lbl === '매출이익') m.grossProfit += total;
            else if (lbl === '영업이익') { m.opProfit += total; m.opProfitData = d; }
            else if (lbl === '손익분기점') m.bep += total;
            else if (lbl === '한계이익') m.marginalProfit += total;
            else if (accKey.startsWith('변동제조비') && accObj.level === 1) m.varMfg += total;
            else if (accKey.startsWith('고정제조비') && accObj.level === 1) m.fixedMfg += total;
            else if ((lbl.includes('판매비') || lbl.includes('일반관리비')) && accObj.level === 1) m.sga += total;

            // 5. 재료비 (원재료비, 부재료비, 직접포장비 등)
            if (lbl.includes('원재료비') || lbl.includes('부재료비') || lbl.includes('직접포장비')) {
                if (accObj.level === 3 || (accObj.level === 2 && !Object.keys(fin).some(k => fin[k].level === 3 && k.startsWith(accObj.key + '_')))) {
                    m.material += total;
                }
            }

            // 6. 통합 노무비 (노무비, 복리비, 복리후생비, 급여, 인건비)
            if (lbl.includes('노무비') || lbl.includes('복리') || lbl.includes('급여') || lbl.includes('인건비')) {
                if (accObj.level === 3 || (accObj.level === 2 && !Object.keys(fin).some(k => fin[k].level === 3 && k.startsWith(accObj.key + '_')))) {
                    m.labor += total;
                }
            }

            // 7. 기타 주요 경비 (외주가공비, 전력비, 소모품비, 포장비, 외주요역비, 일반가공비, 소모공구비, Royalty)
            if (lbl.includes('외주가공비') || lbl.includes('전력비') || lbl.includes('소모품비') || lbl.includes('포장비') || lbl.includes('외주요역비') || lbl.includes('일반가공비') || lbl.includes('소모공구비') || lbl.includes('Royalty') || lbl.includes('로열티')) {
                if (accObj.level === 3 || (accObj.level === 2 && !Object.keys(fin).some(k => fin[k].level === 3 && k.startsWith(accObj.key + '_')))) {
                    m.others += total;
                }
            }

            // 9. 감가상각, 수선, 간접인건비
            if (lbl.includes('감가상각비') || lbl.includes('수선비') || lbl.includes('간접인건비')) {
                if (accObj.level === 3 || (accObj.level === 2 && !Object.keys(fin).some(k => fin[k].level === 3 && k.startsWith(accObj.key + '_')))) {
                    m.deprecMaintInd += total;
                }
            }
        });

        // --- 알고리즘 기반 리스크 감지(Issue Detection) ---
        
        // 1. BEP Overrun (10. 손익분기점 미달)
        if (m.sales > 0 && m.bep > 0 && m.sales < m.bep) {
            dangerCount++;
            issues.push({
                level: 'danger',
                badge: '위험',
                title: `[${selCat}] 10. 손익분기점(BEP) 미달`,
                desc: `조회 기간의 총 매출액(${formatCurr(m.sales)})이 손익분기점(${formatCurr(m.bep)})에 미달하여 구조적인 영업손실이 발생하고 있습니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 변동비 및 고정비 절감 한계치 재평가<br>• 2단계: 손익분기점을 초과할 수 있도록 최소 수주 물량(Q) 확보 및 단가(P) 인상 협상' }
            });
        }

        // 2. Marginal Profit Loss (8. 한계적자 발생)
        if (m.sales > 0 && m.marginalProfit < 0) {
            dangerCount++;
            issues.push({
                level: 'danger',
                badge: '위험',
                title: `[${selCat}] 8. 한계이익 적자 발생`,
                desc: `조회 기간 동안 변동비가 매출을 초과하여 한계이익 적자 ${formatCurr(m.marginalProfit)}를 기록 중입니다. 생산을 늘릴수록 손실이 누적됩니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 부자재 및 물류비 소요 단위 BOM 정밀 재검증 및 단가 긴급 재협상<br>• 2단계: 적자 제품의 수주량을 통제하고 고수익 품목으로 전환 생산' }
            });
        }

        // 3. Gross Profit & Operating Profit Loss (3, 11. 매출이익/영업이익 적자)
        if (m.sales > 0 && m.grossProfit < 0) {
            dangerCount++;
            issues.push({
                level: 'danger',
                badge: '위험',
                title: `[${selCat}] 3. 매출이익 적자 발생`,
                desc: `매출액 대비 제조원가가 초과하여 매출이익이 적자(${formatCurr(m.grossProfit)}) 상태입니다. 제품을 팔수록 원가 손실이 커집니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 직접 재료비 및 노무비 등 핵심 제조원가 낭비 요인 즉각 실사<br>• 2단계: 채산성 한계 제품 단종 및 수익성 위주 포트폴리오 재편' }
            });
        }
        if (m.sales > 0 && m.opProfit < 0 && m.grossProfit >= 0) {
            warningCount++;
            issues.push({
                level: 'warning',
                badge: '주의',
                title: `[${selCat}] 11. 영업이익 적자 발생`,
                desc: `매출이익은 흑자이나, 판관비 등 과다 지출로 인해 영업이익이 적자(${formatCurr(m.opProfit)})를 기록 중입니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 고정 판매비 및 일반관리비 등 간접 경비 제로베이스 예산 통제<br>• 2단계: 조직 슬림화 및 비효율 지원 인력 감축 검토' }
            });
        }

        // 4. Product Valuation Loss Spike (2. 제품평가손실 비중 과다)
        if (m.sales > 0 && m.valLoss > 0 && (m.valLoss / m.sales) > 0.05) {
            warningCount++;
            issues.push({
                level: 'warning',
                badge: '주의',
                title: `[${selCat}] 2. 제품평가손실 비중 과다`,
                desc: `제품평가손실(${formatCurr(m.valLoss)})이 매출의 ${((m.valLoss/m.sales)*100).toFixed(1)}%를 초과하여 재고 자산 건전성이 악화되고 있습니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 장기 체화 재고 및 불량품 현황 전수 조사<br>• 2단계: 재고 덤핑 매각 또는 폐기 처리를 통한 추가 손실 차단' }
            });
        }

        // 5. Labor Cost Burden (6. 통합 노무비 부담 과다)
        if (m.sales > 0 && m.labor > 0 && (m.labor / m.sales) > 0.20) {
            infoCount++;
            issues.push({
                level: 'info',
                badge: '정보',
                title: `[${selCat}] 6. 통합 노무비 부담 과다`,
                desc: `변동/고정/판관비를 통틀어 투입된 총 노무비(복리후생 포함)가 ${formatCurr(m.labor)}으로 매출의 ${((m.labor/m.sales)*100).toFixed(1)}%에 달합니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 공정 라인 자동화율 점검 및 잉여 인력 효율화<br>• 2단계: 잔업/특근 등 변동 인건비 통제 및 유연근무제 도입' }
            });
        }

        // 6. Material Cost Burden (5. 재료비 비중 과다)
        if (m.sales > 0 && m.material > 0 && (m.material / m.sales) > 0.40) {
            infoCount++;
            issues.push({
                level: 'info',
                badge: '정보',
                title: `[${selCat}] 5. 재료비 비중 과다`,
                desc: `핵심 재료비(원재료, 부재료 등)가 ${formatCurr(m.material)}으로 매출의 ${((m.material/m.sales)*100).toFixed(1)}%를 차지하고 있습니다. 원자재가 상승 압력이 큽니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 원부자재 글로벌 소싱 다변화 및 대량 구매 단가 인하 협상<br>• 2단계: 제조 공정 수율(Yield) 개선 및 불량 폐기 감소 활동' }
            });
        }
        
        // 7. Other Major Expenses (7. 8대 기타 주요 경비 집중 관리)
        if (m.sales > 0 && m.others > 0 && (m.others / m.sales) > 0.15) {
            infoCount++;
            issues.push({
                level: 'info',
                badge: '정보',
                title: `[${selCat}] 7. 8대 기타 주요 경비 과다`,
                desc: `외주가공, 전력비, 소모품 등 주요 경비 합산액이 ${formatCurr(m.others)}으로 매출의 ${((m.others/m.sales)*100).toFixed(1)}%를 차지합니다. 추가 절감 여력이 있는지 점검이 필요합니다.`,
                action: { title: '추천 조치 계획', p: '• 1단계: 외주가공 공정의 내재화 검토 및 전력 누수 타임 개선<br>• 2단계: 소모공구 수명 연장 활동 및 비품 구매 승인 절차 강화' }
            });
        }

        // 8. Tet Holiday Impact (2월 조업도 손실) - 매출액 기준
        let startY = parseInt(State.ui.startYear);
        let endY = parseInt(State.ui.endYear);
        for (let y = Math.max(2024, startY); y <= Math.min(2026, endY); y++) {
            let yBase = (y - 2024) * 12;
            let targetIdxJan = yBase + 0 - sIdx;
            let targetIdxFeb = yBase + 1 - sIdx;
            let targetIdxMar = yBase + 2 - sIdx;
            
            if (targetIdxJan >= 0 && targetIdxMar < periodLength && m.salesData.length > 0) {
                let janSales = m.salesData[targetIdxJan] || 0;
                let febSales = m.salesData[targetIdxFeb] || 0;
                let marSales = m.salesData[targetIdxMar] || 0;

                let avgJanMar = (janSales + marSales) / 2;
                if (avgJanMar > 0 && febSales / avgJanMar < 0.75) {
                    let dropPct = ((1 - febSales / avgJanMar) * 100).toFixed(1);
                    warningCount++;
                    issues.push({
                        level: 'warning',
                        badge: '주의',
                        title: `[${selCat}] 1. ${y}년 2월 구정(Tet) 조업도 손실`,
                        desc: `${y}년 2월 매출액은 ${formatCurr(febSales)}로, 전후 월(1,3월) 평균 매출 ${formatCurr(avgJanMar)} 대비 ${dropPct}% 급감하였습니다. 가동 중단에 따른 비조업 손실이 주 원인입니다.`,
                        action: { title: '추천 조치 계획', p: '• 1단계: 구정 전후 집중 생산 추진 및 조기 재고 구축으로 출하 차질 최소화<br>• 2단계: 조기 복귀 인센티브 지원 등을 통해 가동률 복구' }
                    });
                }
            }
        }
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
// ==========================================
// Comparative Analysis (비교 분석) Module
// ==========================================
State.ui.comparativeSelections = [
    { cat: '', startYear: '2026', startMonth: '04', endYear: '2026', endMonth: '04' },
    { cat: '', startYear: '2026', startMonth: '04', endYear: '2026', endMonth: '04' },
    { cat: '', startYear: '2026', startMonth: '04', endYear: '2026', endMonth: '04' },
    { cat: '', startYear: '2026', startMonth: '04', endYear: '2026', endMonth: '04' },
    { cat: '', startYear: '2026', startMonth: '04', endYear: '2026', endMonth: '04' }
];

function initComparativeAnalysis() {
    let btnOpen = document.getElementById('btnOpenComparativeAnalysis');

    if (btnOpen) {
        btnOpen.addEventListener('click', () => {
            switchTab('comparativeAnalysis');
            
            // Only try to render slots and accounts if data is loaded
            if (State.processedData && State.processedData.categories && State.processedData.categories.length > 0) {
                renderCompSlots();
                
                // Default select Sales and OpProfit if not selected
                let hasSales = false, hasOpProfit = false;
                if (!State.unitAnalysis.selectedAccounts) State.unitAnalysis.selectedAccounts = [];
                
                State.unitAnalysis.selectedAccounts.forEach(k => {
                    if (k.includes('매출액') || k.includes('매출단가')) hasSales = true;
                    if (k.includes('영업이익')) hasOpProfit = true;
                });
                
                let tcFin = State.processedData.financials ? (State.processedData.financials[State.processedData.totalCompany] || {}) : {};
                if (!hasSales) {
                    let salesKey = Object.keys(tcFin).find(k => k.includes('매출액'));
                    if(salesKey && !State.unitAnalysis.selectedAccounts.includes(salesKey)) State.unitAnalysis.selectedAccounts.push(salesKey);
                }
                if (!hasOpProfit) {
                    let opKey = Object.keys(tcFin).find(k => k.includes('영업이익'));
                    if(opKey && !State.unitAnalysis.selectedAccounts.includes(opKey)) State.unitAnalysis.selectedAccounts.push(opKey);
                }
                
                // Sync all grids
                [els.uaAccountGrid, els.compAccountGrid, document.getElementById('accountSelector')].forEach(g => {
                    if (g) {
                        g.querySelectorAll('input[type="checkbox"]').forEach(c => {
                            c.checked = State.unitAnalysis.selectedAccounts.includes(c.value);
                        });
                    }
                });
                
                runComparativeAnalysis();
            } else {
                let head = document.getElementById('comparativeTableHead');
                let body = document.getElementById('comparativeTableBody');
                if (head) head.innerHTML = '';
                if (body) body.innerHTML = '<tr><td colspan="10" style="text-align:center; padding: 40px; color:var(--text-secondary);">먼저 엑셀 데이터를 업로드해주세요.</td></tr>';
            }
        });
    }
}

function renderCompSlots() {
    let container = document.getElementById('comparativeTableHead');
    if (!container) return;
    
    let tc = State.processedData.totalCompany;
    let cats = State.processedData.categories;
    let catOptions = `<option value="">-- 품목분류 선택 --</option>`;
    if (cats.includes(tc)) {
        let displayName = tc === '성우비나()' ? '베트남법인' : tc;
        catOptions += `<option value="${tc}">${displayName}</option>`;
    }
    cats.forEach(c => {
        if (c !== tc) {
            catOptions += `<option value="${c}">${formatCategoryLabel(c, false)}</option>`;
        }
    });
    
    let html = `<tr><th style="min-width: 180px; text-align: center; padding: 16px; font-size: 1.15rem; font-weight: 700; color: #fff; vertical-align: middle;">비교 항목 (단가 기준)</th>`;
    for(let i=0; i<5; i++) {
        let mOptions = '';
        for(let m=1; m<=12; m++) {
            let mStr = m.toString().padStart(2, '0');
            mOptions += `<option value="${mStr}">${m}월</option>`;
        }
        
        if (i > 0) html += `<th style="width: 16px; min-width: 16px; padding: 0; border: none; background: transparent;"></th>`;
        html += `
            <th style="min-width: 200px; padding: 16px; vertical-align: top; background: rgba(255,255,255,0.02); border-radius: 6px 6px 0 0;">
                <div class="filter-group" style="text-align: left;">
                    <label style="display: block; margin-bottom: 8px; font-weight: 600; color: var(--accent-cyan); font-size: 0.9rem; text-align: center;">품목 ${i+1}</label>
                    <select class="modern-select comp-cat" data-idx="${i}" style="width: 100%; margin-bottom: 8px;">
                        ${catOptions}
                    </select>
                    <div style="display: flex; flex-direction: column; gap: 4px;">
                        <div style="display: flex; gap: 4px;">
                            <select class="modern-select comp-sy" data-idx="${i}" style="flex: 1; min-width: 0; padding: 6px;">
                                <option value="2024">2024년</option><option value="2025">2025년</option><option value="2026">2026년</option>
                            </select>
                            <select class="modern-select comp-sm" data-idx="${i}" style="flex: 1; min-width: 0; padding: 6px;">${mOptions}</select>
                        </div>
                        <div style="text-align: center; color: rgba(255,255,255,0.85); font-size: 1.25rem; font-weight: 600; padding: 2px 0;">~</div>
                        <div style="display: flex; gap: 4px;">
                            <select class="modern-select comp-ey" data-idx="${i}" style="flex: 1; min-width: 0; padding: 6px;">
                                <option value="2024">2024년</option><option value="2025">2025년</option><option value="2026">2026년</option>
                            </select>
                            <select class="modern-select comp-em" data-idx="${i}" style="flex: 1; min-width: 0; padding: 6px;">${mOptions}</select>
                        </div>
                    </div>
                </div>
            </th>
        `;
    }
    html += `</tr>`;
    container.innerHTML = html;

    // Set initial values
    for(let i=0; i<5; i++) {
        let s = State.ui.comparativeSelections[i];
        let catEl = container.querySelector(`.comp-cat[data-idx="${i}"]`);
        if (catEl && Array.from(catEl.options).some(o => o.value === s.cat)) {
            catEl.value = s.cat;
        }
        container.querySelector(`.comp-sy[data-idx="${i}"]`).value = s.startYear;
        container.querySelector(`.comp-sm[data-idx="${i}"]`).value = s.startMonth;
        container.querySelector(`.comp-ey[data-idx="${i}"]`).value = s.endYear;
        container.querySelector(`.comp-em[data-idx="${i}"]`).value = s.endMonth;
    }

    // Attach listeners
    container.querySelectorAll('select').forEach(sel => {
        sel.addEventListener('change', (e) => {
            let idx = e.target.dataset.idx;
            let targetClass = e.target.className;
            let s = State.ui.comparativeSelections[idx];

            if (targetClass.includes('comp-cat')) s.cat = e.target.value;
            if (targetClass.includes('comp-sy')) s.startYear = e.target.value;
            if (targetClass.includes('comp-sm')) s.startMonth = e.target.value;
            if (targetClass.includes('comp-ey')) s.endYear = e.target.value;
            if (targetClass.includes('comp-em')) s.endMonth = e.target.value;

            // Date validation
            let sIdx = (parseInt(s.startYear) - 2024) * 12 + (parseInt(s.startMonth) - 1);
            let eIdx = (parseInt(s.endYear) - 2024) * 12 + (parseInt(s.endMonth) - 1);
            
            if (sIdx > eIdx) {
                if (targetClass.includes('comp-ey') || targetClass.includes('comp-em')) {
                    s.startYear = s.endYear;
                    s.startMonth = s.endMonth;
                    container.querySelector(`.comp-sy[data-idx="${idx}"]`).value = s.startYear;
                    container.querySelector(`.comp-sm[data-idx="${idx}"]`).value = s.startMonth;
                } else {
                    s.endYear = s.startYear;
                    s.endMonth = s.startMonth;
                    container.querySelector(`.comp-ey[data-idx="${idx}"]`).value = s.endYear;
                    container.querySelector(`.comp-em[data-idx="${idx}"]`).value = s.endMonth;
                }
            }

            runComparativeAnalysis();
        });
    });
}



function runComparativeAnalysis() {
    let head = document.getElementById('comparativeTableHead');
    let body = document.getElementById('comparativeTableBody');
    if (!head || !body) return;

    let activeSlots = State.ui.comparativeSelections.filter(s => s.cat !== '');
    if (activeSlots.length === 0) {
        body.innerHTML = `<tr><td colspan="10" style="text-align:center; padding: 40px; color: var(--accent-red);">최소 1개 이상의 품목분류를 선택해주세요.</td></tr>`;
        return;
    }

    let accountsToCompare = State.unitAnalysis.selectedAccounts ? State.unitAnalysis.selectedAccounts.slice() : [];
    if (accountsToCompare.length === 0) {
        body.innerHTML = `<tr><td colspan="10" style="text-align:center; padding: 40px; color: var(--accent-red);">최소 1개 이상의 계정을 선택해주세요.</td></tr>`;
        return;
    }

    let tbodyHtml = '';
    
    let tcFin = State.processedData.financials[State.processedData.totalCompany] || {};
    let orderedKeys = Object.keys(tcFin);
    accountsToCompare.sort((a, b) => {
        return orderedKeys.indexOf(a) - orderedKeys.indexOf(b);
    });

    let allSlots = State.ui.comparativeSelections; // Always 5 slots

    let slotSalesSums = allSlots.map(slot => {
        if (!slot.cat) return 0;
        let sIdx = (parseInt(slot.startYear) - 2024) * 12 + (parseInt(slot.startMonth) - 1);
        let eIdx = (parseInt(slot.endYear) - 2024) * 12 + (parseInt(slot.endMonth) - 1);
        sIdx = Math.max(0, Math.min(35, sIdx));
        eIdx = Math.max(0, Math.min(35, eIdx));
        let fin = State.processedData.financials[slot.cat];
        if (!fin) return 0;
        let salesKey = Object.keys(fin).find(k => k.includes('매출액'));
        if (salesKey && fin[salesKey]) {
            let d = extractPeriodicData(fin[salesKey].data).slice(sIdx, eIdx + 1);
            return d.reduce((a, b) => a + b, 0);
        }
        return 0;
    });

    accountsToCompare.forEach(accKey => {
        let accObj = tcFin[accKey];
        if (!accObj) return;
        
        let displayLabel = accObj.label;
        if (displayLabel.includes('매출액')) displayLabel = '매출단가';

        tbodyHtml += `<tr style="border-bottom: 1px solid rgba(255,255,255,0.05); transition: background 0.2s;" onmouseover="this.style.background='rgba(255,255,255,0.02)'" onmouseout="this.style.background='transparent'">
            <td style="text-align: left; padding: 16px; font-weight: 600; color: var(--accent-cyan); border-right: 1px solid rgba(255,255,255,0.05);">${displayLabel}</td>`;

        allSlots.forEach((slot, idx) => {
            if (idx > 0) tbodyHtml += `<td style="width: 16px; min-width: 16px; padding: 0; border: none; background: transparent;"></td>`;
            let cellValue = '-';

            if (slot.cat) {
                let sIdx = (parseInt(slot.startYear) - 2024) * 12 + (parseInt(slot.startMonth) - 1);
                let eIdx = (parseInt(slot.endYear) - 2024) * 12 + (parseInt(slot.endMonth) - 1);
                sIdx = Math.max(0, Math.min(35, sIdx));
                eIdx = Math.max(0, Math.min(35, eIdx));

                let fin = State.processedData.financials[slot.cat];
                
                if (fin && fin[accKey]) {
                    let d = extractPeriodicData(fin[accKey].data).slice(sIdx, eIdx + 1);
                    let accSum = d.reduce((a, b) => a + b, 0);

                    let volArr = [];
                    if (State.processedData.volumes[slot.cat] && State.processedData.volumes[slot.cat].reduce((a,b)=>a+b, 0) > 0) {
                        volArr = State.processedData.volumes[slot.cat].slice(sIdx, eIdx + 1);
                    } else if (slot.cat === State.processedData.totalCompany) {
                        volArr = new Array(eIdx - sIdx + 1).fill(0);
                        Object.keys(State.processedData.volumes).forEach(k => {
                            if (k !== State.processedData.totalCompany) {
                                let vArr = State.processedData.volumes[k].slice(sIdx, eIdx + 1);
                                for(let i=0; i<volArr.length; i++) volArr[i] += vArr[i] || 0;
                            }
                        });
                    }
                    let volSum = volArr.reduce((a, b) => a + b, 0);

                    if (volSum > 0) {
                        let unitCost = accSum / volSum;
                        let colorStyle = '';
                        let pctColorStyle = 'color: var(--text-secondary);';
                        
                        if (displayLabel.includes('영업이익')) {
                            let isPositive = unitCost >= 0;
                            colorStyle = isPositive ? 'color: #60a5fa;' : 'color: var(--accent-red);';
                            pctColorStyle = colorStyle;
                        }
                        
                        let pctStr = `<span style="flex: 1; text-align: right; padding-right: 32px;"></span>`;
                        let salesSum = slotSalesSums[idx];
                        if (salesSum && salesSum !== 0) {
                            let pct = (accSum / salesSum) * 100;
                            pctStr = `<span style="flex: 1; text-align: right; padding-right: 32px; font-size: 0.85rem; ${pctColorStyle}">${pct.toFixed(1)}%</span>`;
                        }
                        
                        cellValue = `
                            <div style="display: flex; justify-content: center; align-items: center; width: 100%;">
                                <span style="flex: 1; text-align: right; padding-right: 12px; font-family: 'JetBrains Mono', monospace; font-size: 1.05rem; ${colorStyle}">₩ ${formatUnit(unitCost)}</span>
                                ${pctStr}
                            </div>
                        `;
                    } else if (accSum !== 0) {
                        cellValue = `<span style="color: var(--text-secondary); font-size: 0.85rem;">조업도 없음</span>`;
                    } else {
                        cellValue = `<span style="color: var(--text-secondary); font-family: 'JetBrains Mono', monospace;">₩ 0</span>`;
                    }
                }
            }
            tbodyHtml += `<td style="text-align: right; padding: 16px; background: rgba(255,255,255,0.01);">${cellValue}</td>`;
        });
        tbodyHtml += `</tr>`;
    });

    body.innerHTML = tbodyHtml;
}

document.addEventListener('DOMContentLoaded', init);
