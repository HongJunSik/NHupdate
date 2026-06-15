/* -------------------------------------------------------------
   NHupdate - app.js
   Frontend application logic for Excel parsing, deduplication,
   filtering out foreign teams, and applying user-specified 12 rules
   to format data for easy copy-pasting (29 columns)
------------------------------------------------------------- */

document.addEventListener('DOMContentLoaded', () => {
    // State management
    let state = {
        originalData: [], // Raw rows including header
        cleanedData: [],  // Deduplicated and foreign-filtered rows
        duplicateLog: [], // Duplicated or excluded rows details
        mappedData: [],   // Mapped to Google Sheets 29 columns
        fileName: ''
    };

    // DOM Elements
    const btnDashboard = document.getElementById('btn-dashboard');
    const btnMapping = document.getElementById('btn-mapping');
    
    const sectionDashboard = document.getElementById('section-dashboard');
    const sectionMapping = document.getElementById('section-mapping');
    
    const pageTitle = document.getElementById('page-title');
    const pageDesc = document.getElementById('page-desc');
    
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    
    const summarySection = document.getElementById('summary-section');
    const loadedFileName = document.getElementById('loaded-file-name');
    const originalRowCount = document.getElementById('original-row-count');
    const cleanedRowCount = document.getElementById('cleaned-row-count');
    const duplicateRemovedCount = document.getElementById('duplicate-removed-count');
    
    const dataPanel = document.getElementById('data-panel');
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    
    const tableCleaned = document.getElementById('table-cleaned');
    const tableOriginal = document.getElementById('table-original');
    const dupeLogList = document.getElementById('dupe-log-list');
    const dupeTabCount = document.getElementById('dupe-tab-count');
    
    const btnDownloadCleaned = document.getElementById('btn-download-cleaned');
    const btnCopyClipboard = document.getElementById('btn-copy-clipboard');
    const btnCopyClipboardBanner = document.getElementById('btn-copy-clipboard-banner');
    const actionBanner = document.getElementById('action-banner');
    
    const toast = document.getElementById('toast');

    // 1. Navigation Logic
    const sections = [
        { btn: btnDashboard, sect: sectionDashboard, title: '복사 대시보드', desc: '당일 예약조회 엑셀 파일을 첨부하면 내원현황표 열 규격(29열)에 맞춰 복사용 데이터가 생성됩니다.' },
        { btn: btnMapping, sect: sectionMapping, title: '지정 매핑 규칙', desc: '엑셀 파일의 데이터가 구글 내원현황표 시트의 어떤 항목으로 가공되어 들어가는지 설명합니다.' }
    ];

    sections.forEach(item => {
        item.btn.addEventListener('click', () => {
            sections.forEach(s => {
                s.btn.classList.remove('active');
                s.sect.classList.remove('active');
            });
            item.btn.classList.add('active');
            item.sect.classList.add('active');
            pageTitle.textContent = item.title;
            pageDesc.textContent = item.desc;
        });
    });

    // Set Current Date
    const today = new Date();
    const formattedToday = `${today.getFullYear()}. ${String(today.getMonth() + 1).padStart(2, '0')}. ${String(today.getDate()).padStart(2, '0')}`;
    document.getElementById('current-date').textContent = formattedToday;

    // 2. Drag & Drop Upload Handlers
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
    
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
            fileInput.value = ''; // Reset input value to allow uploading the same or another file again
        }
    });

    function resetUI() {
        summarySection.style.display = 'none';
        actionBanner.style.display = 'none';
        dataPanel.style.display = 'none';
    }

    // 3. File Processing Logic
    function handleFile(file) {
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showToast("올바른 엑셀 파일(.xlsx)을 업로드해주세요.", "error");
            return;
        }
        
        state.fileName = file.name;
        
        // Reset UI before processing
        resetUI();
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Get raw 2D array of values
                const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (rawRows.length < 2) {
                    showToast("엑셀 파일에 데이터가 부족합니다.", "error");
                    return;
                }
                
                processExcelData(rawRows);
                showToast("파일을 성공적으로 가공했습니다.");
            } catch (err) {
                console.error(err);
                showToast(err.message || "파일 파싱 중 에러가 발생했습니다.", "error");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Data Processing Helper Functions
    function extractAge(memo) {
        if (!memo) return '';
        const dateMatch = memo.match(/([12]\d{3})[\/\.\-]\d{1,2}[\/\.\-]\d{1,2}/);
        if (dateMatch) {
            const birthYear = parseInt(dateMatch[1]);
            const currentYear = 2026; 
            return currentYear - birthYear;
        }
        const yearMatch = memo.match(/(?:19|20)\d{2}/);
        if (yearMatch) {
            const birthYear = parseInt(yearMatch[0]);
            const currentYear = 2026;
            return currentYear - birthYear;
        }
        return '';
    }

    function extractGender(memo) {
        if (!memo) return '';
        if (/여성|여자|\b여\b|female|f\b/i.test(memo)) {
            return 'F';
        }
        if (/남성|남자|\b남\b|male|m\b/i.test(memo)) {
            return 'M';
        }
        return '';
    }

    function formatDate(yyyymmdd) {
        if (!yyyymmdd) return '';
        const str = String(yyyymmdd).trim();
        if (str.length === 8 && /^\d+$/.test(str)) {
            const month = parseInt(str.substring(4, 6), 10);
            const day = parseInt(str.substring(6, 8), 10);
            return `${month}월 ${day}일`;
        }
        // Match YYYY-MM-DD
        const ymdMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
        if (ymdMatch) {
            return `${parseInt(ymdMatch[2], 10)}월 ${parseInt(ymdMatch[3], 10)}일`;
        }
        // Match MM-DD or M-D
        const mdMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})$/);
        if (mdMatch) {
            return `${parseInt(mdMatch[1], 10)}월 ${parseInt(mdMatch[2], 10)}일`;
        }
        return yyyymmdd;
    }

    function formatTime(timeStr) {
        if (!timeStr) return '';
        const str = String(timeStr).trim();
        if (str.length === 4 && /^\d+$/.test(str)) {
            const hour = str.substring(0, 2);
            const min = str.substring(2, 4);
            return `${hour}:${min}`;
        }
        return timeStr;
    }

    function processExcelData(rawRows) {
        // Filter out completely empty rows
        const validRows = rawRows.filter(row => row && row.length > 0 && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''));
        
        state.originalData = validRows;
        
        const headers = validRows[0];
        
        // Find column indices
        const colIdx = {};
        headers.forEach((h, idx) => {
            if (h) colIdx[String(h).trim()] = idx;
        });

        // Check if required columns exist
        const required = ['고객번호', '고객명', '예약일', '예약시간', '진료실', '고객메모', '예약상태', '고객구분', '성별', '나이', '예약메모', '기타메모'];
        const missing = required.filter(col => colIdx[col] === undefined);
        if (missing.length > 0) {
            throw new Error(`해당 파일에는 '${missing[0]}' 데이터가 없습니다.`);
        }

        const uniqueCustomers = new Map();
        const exclusions = []; // To keep logs of duplicates AND excluded foreign teams
        const cleaned = [headers]; // include header row
        
        const custNumIdx = colIdx['고객번호'];
        const custNameIdx = colIdx['고객명'];
        const timeIdx = colIdx['예약시간'];
        const memoIdx = colIdx['고객메모'];
        const statusIdx = colIdx['예약상태'] || -1;
        
        for (let i = 1; i < validRows.length; i++) {
            const row = validRows[i];
            const custNum = String(row[custNumIdx] || '').trim();
            const custName = String(row[custNameIdx] || '').trim();
            const time = String(row[timeIdx] || '').trim();
            const status = statusIdx !== -1 ? String(row[statusIdx] || '').trim() : '';
            const custMemo = String(row[memoIdx] || '').trim();
            
            // 규칙 1-2: 고객메모에 '일본팀', '대만팀', '해외' 등이 들어가면 제외(삭제)
            if (custMemo.includes('일본') || custMemo.includes('대만') || custMemo.includes('해외')) {
                exclusions.push({
                    type: 'foreign',
                    customerNumber: custNum,
                    name: custName,
                    time: time,
                    status: status,
                    reason: `해외팀 제외 (${custMemo.split('\n')[0]})`,
                    rowNum: i + 1
                });
                continue; // Skip writing this row to cleaned list
            }
            
            // 규칙 1-3: 고객명에 '예약' 또는 '마감'이 들어가면 제외(삭제)
            if (custName.includes('예약') || custName.includes('마감')) {
                exclusions.push({
                    type: 'reference',
                    customerNumber: custNum,
                    name: custName,
                    time: time,
                    status: status,
                    reason: '예약 시 참고용',
                    rowNum: i + 1
                });
                continue; // Skip writing this row to cleaned list
            }
            
            if (!custNum) {
                cleaned.push(row);
                continue;
            }
            
            // 규칙 1: 고객번호 비교하여 중복 고객은 1행만 보존
            if (uniqueCustomers.has(custNum)) {
                // Duplicate booking found
                exclusions.push({
                    type: 'duplicate',
                    customerNumber: custNum,
                    name: custName,
                    time: time,
                    status: status,
                    reason: '중복 고객 예약 제거',
                    rowNum: i + 1
                });
                
                const mainRecord = uniqueCustomers.get(custNum);
                mainRecord.duplicates.push({
                    rowNum: i + 1,
                    time: time,
                    status: status
                });
            } else {
                // First reservation of this customer
                const record = {
                    rowNum: i + 1,
                    name: custName,
                    time: time,
                    status: status,
                    row: row,
                    duplicates: []
                };
                uniqueCustomers.set(custNum, record);
                cleaned.push(row);
            }
        }
        
        state.cleanedData = cleaned;
        state.duplicateLog = exclusions;

        // Build Mapped Data for Google Sheet 29-column structure
        state.mappedData = Array.from(uniqueCustomers.values()).map(item => {
            const r = item.row;
            const memo = r[colIdx['예약메모']] || '';
            const custMemo = r[colIdx['고객메모']] || '';
            const otherMemo = r[colIdx['기타메모']] || '';
            const rawDate = r[colIdx['예약일']] || '';
            
            // 규칙 10: 성별/나이 열이 엑셀에 존재하면 사용, 없으면 메모 파싱
            const gender = colIdx['성별'] !== undefined ? (r[colIdx['성별']] || '') : extractGender(memo);
            const age = colIdx['나이'] !== undefined ? (r[colIdx['나이']] || '') : extractAge(memo);
            
            // 규칙 2, 10: 예약날짜, 예약시간 포맷 변환
            const formattedDate = formatDate(rawDate);
            const formattedTime = formatTime(r[colIdx['예약시간']] || '');

            return {
                reservationDate: formattedDate,                   // 1: 예약 날짜 (규칙 2)
                registrant: '홍준식',                              // 2: 확인자 (규칙 3 - '홍준식' 고정)
                customerName: r[colIdx['고객명']] || '',           // 3: 이름 (규칙 4 - 고객명)
                visitPurpose: r[colIdx['진료실']] || '',           // 4: 내원 목적 (규칙 5 - 진료실)
                marketingPath: r[colIdx['고객메모']] || '',          // 5: 내원경로 (규칙 6 - 고객메모)
                doctor: r[colIdx['예약상태']] || '',               // 6: 담당의사 (규칙 7 - 예약상태)
                coordinator: '',                                   // 7: 상담실장 (규칙 8 - 비워둠)
                patientType: r[colIdx['고객구분']] || '',           // 8: 초/재진 (규칙 9 - 고객구분)
                gender: gender,                                    // 9: 성별 (규칙 10)
                age: age,                                          // 10: 나이 (규칙 10)
                reservationTime: formattedTime,                    // 11: 예약시간 (규칙 10)
                // 12~20: 내원시간 ~ 상담내용 (규칙 11 - 비워둠)
                memo: memo,                                        // 21: 전달사항 (규칙 12 - 예약메모)
                customerMemo: otherMemo                            // 22: 특이사항 (기타메모 매핑)
            };
        });

        // Update UI Summary
        loadedFileName.textContent = state.fileName;
        originalRowCount.textContent = `${validRows.length - 1}개`;
        cleanedRowCount.textContent = `${state.mappedData.length}개`;
        duplicateRemovedCount.textContent = `${exclusions.length}개`;
        dupeTabCount.textContent = exclusions.length;
        
        summarySection.style.display = 'grid';
        actionBanner.style.display = 'block';
        dataPanel.style.display = 'flex';
        
        // Render Tables
        renderTable(tableOriginal, state.originalData);
        renderMappedTable(tableCleaned, state.mappedData);
        renderDuplicateLog(uniqueCustomers, exclusions);
    }

    function renderTable(tableElem, rowList) {
        const thead = tableElem.querySelector('thead');
        const tbody = tableElem.querySelector('tbody');
        
        thead.innerHTML = '';
        tbody.innerHTML = '';
        
        if (rowList.length === 0) return;
        
        const headers = rowList[0];
        const trHead = document.createElement('tr');
        headers.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h || '';
            trHead.appendChild(th);
        });
        thead.appendChild(trHead);
        
        const displayLimit = Math.min(rowList.length, 100);
        for (let i = 1; i < displayLimit; i++) {
            const row = rowList[i];
            const tr = document.createElement('tr');
            headers.forEach((h, idx) => {
                const td = document.createElement('td');
                td.textContent = row[idx] !== undefined && row[idx] !== null ? row[idx] : '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        }
    }

    function renderMappedTable(tableElem, mappedList) {
        const thead = tableElem.querySelector('thead');
        const tbody = tableElem.querySelector('tbody');
        
        thead.innerHTML = '';
        tbody.innerHTML = '';
        
        const cols = [
            "예약 날짜", "확인자", "이름", "내원 목적", "내원경로", "담당의사", "상담실장", "초/재진", 
            "성별", "나이", "예약시간", "내원시간", "상담시작", "상담종료", "입실시간", "퇴원시간", 
            "상담대기", "시술대기", "15분대기사유", "상담내용", "전달사항", "특이사항", "과세", "비과세", 
            "누적매출", "평균일매출", "목표 일매출", "예상매출액", "주차별"
        ];

        // Headers
        const trHead = document.createElement('tr');
        cols.forEach(label => {
            const th = document.createElement('th');
            th.textContent = label;
            trHead.appendChild(th);
        });
        thead.appendChild(trHead);

        // Body
        mappedList.forEach(item => {
            const tr = document.createElement('tr');
            
            // Build the exact 29 columns row based on rules
            const rowValues = [
                item.reservationDate,      // 1: 예약 날짜
                item.registrant,           // 2: 확인자 ('홍준식' 고정)
                item.customerName,         // 3: 이름 (고객명)
                item.visitPurpose,         // 4: 내원 목적 (진료실)
                item.marketingPath,        // 5: 내원경로 (고객메모)
                item.doctor,               // 6: 담당의사 (예약상태)
                item.coordinator,          // 7: 상담실장 (비워둠)
                item.patientType,          // 8: 초/재진 (고객구분)
                item.gender,               // 9: 성별
                item.age,                  // 10: 나이
                item.reservationTime,      // 11: 예약시간
                // 12~20: 내원시간 ~ 상담내용 (비워둠)
                "", "", "", "", "", "", "", "", "", 
                item.memo,                 // 21: 전달사항 (예약메모)
                "",                        // 22: 특이사항 (비워둠)
                "",                        // 23: 과세 (비워둠)
                "",                        // 24: 비과세 (비워둠)
                // 25~29: 누적매출 ~ 주차별 (비워둠)
                "", "", "", "", ""         
            ];

            rowValues.forEach((val, idx) => {
                const td = document.createElement('td');
                td.textContent = val !== undefined && val !== null ? val : '';
                
                // Highlight key columns
                if (idx === 2 || idx === 0 || idx === 10) { // 이름, 예약 날짜, 예약시간
                    td.classList.add('highlight');
                }
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }

    function renderDuplicateLog(uniqueMap, exclusionsList) {
        dupeLogList.innerHTML = '';
        
        if (exclusionsList.length === 0) {
            dupeLogList.innerHTML = '<div class="no-dupes-msg" style="text-align: center; color: var(--text-dark); padding: 40px;">제외된 중복 또는 해외 예약 건이 없습니다. ✨</div>';
            return;
        }

        exclusionsList.forEach(item => {
            const dupeGroup = document.createElement('div');
            dupeGroup.classList.add('dupe-group');
            
            // Set style and label depending on type
            let badgeLabel = '중복';
            let badgeColor = 'var(--accent-red)';
            let badgeBg = 'rgba(239, 68, 68, 0.15)';
            
            if (item.type === 'foreign') {
                dupeGroup.style.borderColor = 'rgba(245, 158, 11, 0.2)';
                dupeGroup.style.backgroundColor = 'rgba(245, 158, 11, 0.02)';
                badgeLabel = '해외팀 예약';
                badgeColor = 'var(--accent-orange)';
                badgeBg = 'rgba(245, 158, 11, 0.15)';
            } else if (item.type === 'reference') {
                dupeGroup.style.borderColor = 'rgba(168, 85, 247, 0.2)';
                dupeGroup.style.backgroundColor = 'rgba(168, 85, 247, 0.02)';
                badgeLabel = '예약 시 참고용';
                badgeColor = 'var(--accent-purple)';
                badgeBg = 'rgba(168, 85, 247, 0.15)';
            }
            
            dupeGroup.innerHTML = `
                <div class="dupe-header" style="margin-bottom: 0; display: flex; align-items: center; justify-content: space-between; width: 100%;">
                    <div class="dupe-header-title">
                        행 <strong>${item.rowNum}</strong> | 성함: <strong>${item.name}</strong> (고객번호: <span>${item.customerNumber || '없음'}</span>)
                        <span style="color: var(--text-muted); font-size: 12px; margin-left: 12px;">시간: ${formatTime(item.time)} | 상태: ${item.status}</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <span style="font-size: 12px; color: var(--text-muted);">제외 사유:</span>
                        <div class="dupe-badge" style="background-color: ${badgeBg}; color: ${badgeColor}; font-weight: 700; border-radius: 4px; padding: 4px 10px;">${badgeLabel}</div>
                    </div>
                </div>
            `;
            dupeLogList.appendChild(dupeGroup);
        });
    }

    // Tab buttons handler
    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            tabButtons.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));
            
            btn.classList.add('active');
            const targetTab = btn.getAttribute('data-tab');
            document.getElementById(targetTab).classList.add('active');
        });
    });

    // 4. TSV Field Escaper (to protect newlines/tabs inside cells when copying)
    function escapeTSVField(val) {
        if (val === undefined || val === null) return '';
        let str = String(val);
        if (str.includes('\t') || str.includes('\n') || str.includes('\r') || str.includes('"')) {
            str = str.replace(/"/g, '""');
            return `"${str}"`;
        }
        return str;
    }

    // Fallback clipboard function for file:// (non-secure context)
    function copyTextToClipboard(text) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            return navigator.clipboard.writeText(text);
        }
        return new Promise((resolve, reject) => {
            try {
                const textArea = document.createElement("textarea");
                textArea.value = text;
                textArea.style.top = "0";
                textArea.style.left = "0";
                textArea.style.position = "fixed";
                textArea.style.opacity = "0";
                document.body.appendChild(textArea);
                textArea.focus();
                textArea.select();
                const successful = document.execCommand('copy');
                document.body.removeChild(textArea);
                if (successful) {
                    resolve();
                } else {
                    reject(new Error("클립보드 복사 명령(execCommand)이 거부되었습니다."));
                }
            } catch (err) {
                reject(err);
            }
        });
    }

    // 5. Copy to Clipboard (TSV format for direct Sheets pasting)
    function copyToClipboard() {
        if (state.mappedData.length === 0) {
            showToast("복사할 데이터가 없습니다.", "error");
            return;
        }
        
        // Generate TSV rows (without headers, so they paste directly into existing sheet rows)
        const tsvRows = state.mappedData.map(item => {
            const row = [
                item.reservationDate,      // 1: 예약 날짜 (예: 7월 1일)
                item.registrant,           // 2: 확인자 ('홍준식' 고정)
                item.customerName,         // 3: 이름 (고객명)
                item.visitPurpose,         // 4: 내원 목적 (진료실)
                item.marketingPath,        // 5: 내원경로 (고객메모)
                item.doctor,               // 6: 담당의사 (예약상태)
                item.coordinator,          // 7: 상담실장 (비워둠)
                item.patientType,          // 8: 초/재진 (고객구분)
                item.gender,               // 9: 성별
                item.age,                  // 10: 나이
                item.reservationTime,      // 11: 예약시간
                // 12~20: 내원시간 ~ 상담내용 (비워둠)
                "", "", "", "", "", "", "", "", "", 
                item.memo,                 // 21: 전달사항 (예약메모)
                "",                        // 22: 특이사항 (비워둠)
                "",                        // 23: 과세 (비워둠)
                "",                        // 24: 비과세 (비워둠)
                // 25~29: 누적매출 ~ 주차별 (비워둠)
                "", "", "", "", ""
            ];
            return row.map(escapeTSVField).join('\t');
        });
        
        const tsvContent = tsvRows.join('\n');
        
        copyTextToClipboard(tsvContent).then(() => {
            showToast(`내원현황표 복사용 데이터(${state.mappedData.length}건, 29열)가 복사되었습니다! 구글 시트의 A열에 붙여넣으세요.`);
        }).catch(err => {
            console.error("클립보드 복사 실패", err);
            showToast("클립보드 복사에 실패했습니다.", "error");
        });
    }

    btnCopyClipboard.addEventListener('click', copyToClipboard);
    btnCopyClipboardBanner.addEventListener('click', copyToClipboard);

    // 6. Download Cleaned Excel File (29-column structure)
    btnDownloadCleaned.addEventListener('click', () => {
        if (state.mappedData.length === 0) return;
        
        const wb = XLSX.utils.book_new();
        
        // 29-column headers
        const headers = [
            "예약 날짜", "확인자", "이름", "내원 목적", "내원경로", "담당의사", "상담실장", "초/재진", 
            "성별", "나이", "예약시간", "내원시간", "상담시작", "상담종료", "입실시간", "퇴원시간", 
            "상담대기", "시술대기", "15분대기사유", "상담내용", "전달사항", "특이사항", "과세", "비과세", 
            "누적매출", "평균일매출", "목표 일매출", "예상매출액", "주차별"
        ];
        
        const rows = state.mappedData.map(item => [
            item.reservationDate,
            item.registrant,
            item.customerName,
            item.visitPurpose,
            item.marketingPath,
            item.doctor,
            item.coordinator,
            item.patientType,
            item.gender,
            item.age,
            item.reservationTime,
            "", "", "", "", "", "", "", "", "", // 12~20
            item.memo,
            "",
            "",
            "",
            "", "", "", "", "" // 25~29
        ]);
        
        rows.unshift(headers);
        
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, "내원현황표_복사용");
        
        const cleanName = state.fileName.replace('.xlsx', '_정제완료.xlsx').replace('.xls', '_정제완료.xlsx');
        XLSX.writeFile(wb, cleanName);
        showToast("정제 완료된 엑셀 파일이 다운로드되었습니다.");
    });

    // Toast utility
    function showToast(message, type = 'success') {
        console.log("showToast called with message:", message, "type:", type);
        toast.className = 'toast';
        
        if (type === 'error') {
            toast.classList.add('error');
            toast.querySelector('.toast-icon').textContent = '✕';
            toast.style.borderColor = 'var(--accent-red)';
            toast.querySelector('.toast-icon').style.color = 'var(--accent-red)';
        } else {
            toast.querySelector('.toast-icon').textContent = '✓';
            toast.style.borderColor = 'var(--accent-cyan)';
            toast.querySelector('.toast-icon').style.color = 'var(--accent-cyan)';
        }
        
        toast.querySelector('.toast-message').textContent = message;
        toast.classList.add('show');
        
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }
});
