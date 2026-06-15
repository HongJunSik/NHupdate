const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const XLSX = require('xlsx');

// 1. Load HTML and create JSDOM instance
const htmlPath = path.join(__dirname, 'index.html');
const cssPath = path.join(__dirname, 'styles.css');
const jsPath = path.join(__dirname, 'app.js');

const html = fs.readFileSync(htmlPath, 'utf8');
const css = fs.readFileSync(cssPath, 'utf8');

const jsdom = require('jsdom');
const virtualConsole = new jsdom.VirtualConsole();
virtualConsole.on("log", (...args) => console.log("[Virtual Console Log]", ...args));
virtualConsole.on("error", (...args) => console.error("[Virtual Console Error]", ...args));
virtualConsole.on("warn", (...args) => console.warn("[Virtual Console Warn]", ...args));
virtualConsole.on("jsdomError", (err) => console.error("[JSDOM Error]", err.message, err.stack));

const dom = new JSDOM(html, {
    runScripts: "dangerously",
    virtualConsole: virtualConsole
});

const { window } = dom;
const { document } = window;

// 2. Attach mock libraries and globals
window.XLSX = XLSX;

// Mock FileReader
class MockFileReader {
    readAsArrayBuffer(file) {
        console.log("MockFileReader.readAsArrayBuffer called for:", file.name);
        const filePath = path.join(__dirname, file.name);
        try {
            console.log("Reading file from path:", filePath);
            const buffer = fs.readFileSync(filePath);
            console.log("File read successfully, buffer length:", buffer.length);
            this.result = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
            if (this.onload) {
                console.log("Triggering onload callback...");
                this.onload({ target: { result: this.result } });
            }
        } catch (e) {
            console.error("Error in MockFileReader:", e);
            if (this.onerror) this.onerror(e);
        }
    }
}
window.FileReader = MockFileReader;

// Mock Clipboard
let clipboardContent = '';
window.navigator.clipboard = {
    writeText: (text) => {
        clipboardContent = text;
        return Promise.resolve();
    }
};

// Mock alert/toast showToast function is wired inside DOM
// Let's load app.js in jsdom context
const jsCode = fs.readFileSync(jsPath, 'utf8');
const scriptElement = document.createElement("script");
scriptElement.textContent = jsCode;
document.body.appendChild(scriptElement);

// 3. Trigger DOMContentLoaded event
const event = new window.Event('DOMContentLoaded');
document.dispatchEvent(event);

// 4. Run Test Suite
async function runTest() {
    console.log("=== Starting Web App Integration Test ===");
    
    // Check initial state
    // (connection-status is removed in simplified index.html)
    
    // Simulate drop/upload of '예약조회 (1).xlsx'
    const fileInput = document.getElementById('file-input');
    const mockFile = { name: '예약조회 (1).xlsx' };
    
    // We simulate the change event on the file input
    const changeEvent = new window.Event('change');
    
    // We override 'files' directly on the fileInput element
    Object.defineProperty(fileInput, 'files', {
        value: [mockFile],
        configurable: true
    });
    
    console.log("Simulating file upload of '예약조회 (1).xlsx'...");
    fileInput.dispatchEvent(changeEvent);
    
    // Wait for async FileReader callback (it runs synchronously in our mock, but let's wait a bit)
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // Verify KPI Summary values
    const originalCountVal = document.getElementById('original-row-count').textContent;
    const cleanedCountVal = document.getElementById('cleaned-row-count').textContent;
    const removedCountVal = document.getElementById('duplicate-removed-count').textContent;
    const loadedNameVal = document.getElementById('loaded-file-name').textContent;
    
    console.log("\n--- KPI Summary Verification ---");
    console.log("Loaded File Name:", loadedNameVal, "(Expected: 예약조회 (1).xlsx)");
    console.log("Original Row Count:", originalCountVal, "(Expected: 84개)");
    console.log("Cleaned Row Count:", cleanedCountVal, "(Expected: 55개)");
    console.log("Removed/Excluded Count:", removedCountVal, "(Expected: 29개)");
    
    if (originalCountVal !== "84개" || cleanedCountVal !== "55개" || removedCountVal !== "29개") {
        console.error("FAIL: KPI Counts do not match expected values!");
        process.exit(1);
    }
    console.log("SUCCESS: KPI Counts verified!");

    // Verify Tab Data
    console.log("\n--- DOM Table Layout Verification ---");
    const cleanedTable = document.getElementById('table-cleaned');
    const headers = Array.from(cleanedTable.querySelectorAll('thead th')).map(th => th.textContent);
    console.log("Cleaned Table Columns Count:", headers.length, "(Expected: 29 columns)");
    
    if (headers.length !== 29) {
        console.error("FAIL: Table does not have exactly 29 columns!");
        process.exit(1);
    }
    console.log("Columns match 29-column Visit Status Sheet structure!");
    
    // Check first data row names and cells
    const firstRowCells = Array.from(cleanedTable.querySelectorAll('tbody tr')[0].querySelectorAll('td')).map(td => td.textContent);
    console.log("First row name:", firstRowCells[2], "(Expected: 최은영)");
    console.log("First row registrant:", firstRowCells[1], "(Expected: 홍준식)");
    console.log("First row gender:", firstRowCells[8], "(Expected: F)");
    console.log("First row age:", firstRowCells[9], "(Expected: 48)");
    console.log("First row time:", firstRowCells[10], "(Expected: 09:00)");
    
    if (firstRowCells[1] !== '홍준식' || firstRowCells[2] !== '최은영' || firstRowCells[8] !== 'F' || firstRowCells[9] !== '48') {
        console.error("FAIL: First data row cell mapping is incorrect!");
        process.exit(1);
    }
    console.log("SUCCESS: Mapped cell values are correct!");

    // Verify Duplicate Log Exclusions reasons
    console.log("\n--- Duplicate/Exclusion Log Verification ---");
    const dupeLogs = Array.from(document.getElementById('dupe-log-list').querySelectorAll('.dupe-group'));
    console.log("Total exclusions logged:", dupeLogs.length, "(Expected: 29)");
    
    const sampleDupeLog = dupeLogs[0].innerHTML;
    console.log("Sample exclusion log content snippet:");
    console.log("  " + dupeLogs[0].querySelector('.dupe-header-title').textContent.trim());
    console.log("  Exclusion Badge:", dupeLogs[0].querySelector('.dupe-badge').textContent.trim());
    
    // Check if reasons contain '해외팀 예약', '중복' and '예약 시 참고용'
    let hasForeignReason = false;
    let hasDuplicateReason = false;
    let hasReferenceReason = false;
    dupeLogs.forEach(log => {
        const text = log.querySelector('.dupe-badge').textContent;
        if (text.includes('해외팀 예약')) hasForeignReason = true;
        if (text.includes('중복')) hasDuplicateReason = true;
        if (text.includes('예약 시 참고용')) hasReferenceReason = true;
    });
    
    console.log("Exclusion reasons verified: '해외팀 예약' exists?", hasForeignReason, "| '중복' exists?", hasDuplicateReason, "| '예약 시 참고용' exists?", hasReferenceReason);
    if (!hasForeignReason || !hasDuplicateReason || !hasReferenceReason) {
        console.error("FAIL: Exclusion reasons did not contain expected labels!");
        process.exit(1);
    }
    console.log("SUCCESS: Log labels verified!");

    // Verify Copy to Clipboard Button
    console.log("\n--- Clipboard Copy Functionality Verification ---");
    const copyBtn = document.getElementById('btn-copy-clipboard');
    const copyBtnBanner = document.getElementById('btn-copy-clipboard-banner');
    
    console.log("Triggering copy event via the main header button...");
    copyBtn.click();
    
    // Helper to parse TSV with quoted fields containing newlines
    function parseTSV(text) {
        const rows = [];
        let currentRow = [];
        let currentCell = '';
        let insideQuotes = false;
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            const nextChar = text[i+1];
            if (char === '"') {
                if (insideQuotes && nextChar === '"') {
                    currentCell += '"';
                    i++;
                } else {
                    insideQuotes = !insideQuotes;
                }
            } else if (char === '\t' && !insideQuotes) {
                currentRow.push(currentCell);
                currentCell = '';
            } else if ((char === '\n' || char === '\r') && !insideQuotes) {
                if (char === '\r' && nextChar === '\n') i++;
                currentRow.push(currentCell);
                rows.push(currentRow);
                currentRow = [];
                currentCell = '';
            } else {
                currentCell += char;
            }
        }
        if (currentCell || currentRow.length > 0) {
            currentRow.push(currentCell);
            rows.push(currentRow);
        }
        return rows;
    }
    
    const parsedRows = parseTSV(clipboardContent.trim());
    console.log("Copied rows count:", parsedRows.length, "(Expected: 55)");
    const colsInRow = parsedRows[0];
    console.log("Columns count in copied TSV row:", colsInRow.length, "(Expected: 29)");
    
    if (parsedRows.length !== 55 || colsInRow.length !== 29) {
        console.error("FAIL: Copied clipboard content structure is incorrect!");
        process.exit(1);
    }
    console.log("SUCCESS: Clipboard TSV output is correct!");
    
    // Verify fallback clipboard copy
    console.log("Testing clipboard fallback in non-secure context...");
    // Force navigator.clipboard to be undefined
    delete window.navigator.clipboard;
    
    // Mock document.execCommand
    let execCopyTriggered = false;
    document.execCommand = (command) => {
        if (command === 'copy') {
            execCopyTriggered = true;
            return true;
        }
        return false;
    };
    
    console.log("Triggering copy event via the action banner button...");
    copyBtnBanner.click();
    console.log("execCommand('copy') triggered?", execCopyTriggered);
    
    if (!execCopyTriggered) {
        console.error("FAIL: Fallback copy mechanism did not trigger document.execCommand!");
        process.exit(1);
    }
    console.log("SUCCESS: Clipboard fallback verified!");
    
    // Wait for copy promises to settle and toasts to display
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Verify Column Validation Failure
    console.log("\n--- Required Column Validation Failure Verification ---");
    const invalidFilePath = path.join(__dirname, '예약조회_오류테스트.xlsx');
    try {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([
            ['Checked', 'No', '고객명', '예약상태', '고객메모', '담당의사', '예약일', '고객구분', '나이', '예약시간', '예약메모', '진료실', '고객번호', '상담사', '기타메모'], // missing '성별'
            ['', '', '테스트', '', '', '', '', '', '', '', '', '', '', '', ''] // dummy data row
        ]);
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, invalidFilePath);
    } catch (err) {
        console.error("Failed to create temporary invalid Excel file:", err);
        process.exit(1);
    }
    
    const invalidFile = { name: '예약조회_오류테스트.xlsx' };
    Object.defineProperty(fileInput, 'files', {
        value: [invalidFile],
        configurable: true
    });
    
    console.log("Simulating upload of file with missing '성별' column...");
    fileInput.dispatchEvent(changeEvent);
    
    await new Promise(resolve => setTimeout(resolve, 500));
    
    const toast = document.getElementById('toast');
    console.log("Toast outerHTML:", toast.outerHTML);
    const toastMsg = toast.querySelector('.toast-message').textContent;
    const isDashboardHidden = document.getElementById('summary-section').style.display === 'none';
    
    console.log("Toast error message:", toastMsg, "(Expected: 해당 파일에는 '성별' 데이터가 없습니다.)");
    console.log("Is dashboard hidden?", isDashboardHidden, "(Expected: true)");
    
    // Clean up temporary file
    try {
        fs.unlinkSync(invalidFilePath);
    } catch (e) {
        console.error("Cleanup error:", e);
    }
    
    if (toastMsg !== "해당 파일에는 '성별' 데이터가 없습니다." || !isDashboardHidden) {
        console.error("FAIL: Required column validation did not reject processing or reset UI correctly!");
        process.exit(1);
    }
    console.log("SUCCESS: Missing column validation verified!");

    // Verify CSS styles for scrollbars and layout
    console.log("\n--- CSS Style Rules Verification ---");
    // Verify .main-content properties in CSS
    const mainContentRegex = /\.main-content\s*\{[^}]*min-width:\s*0/i;
    const tableScrollRegex = /\.table-scroll\s*\{[^}]*overflow-x:\s*scroll\s*!important/i;
    
    const mainContentOk = mainContentRegex.test(css);
    const tableScrollOk = tableScrollRegex.test(css);
    
    console.log(".main-content flexbox min-width:0 fix exists?", mainContentOk);
    console.log(".table-scroll overflow-x:scroll !important exists?", tableScrollOk);
    
    if (!mainContentOk || !tableScrollOk) {
        console.error("FAIL: Layout CSS fixes are missing in styles.css!");
        process.exit(1);
    }
    console.log("SUCCESS: CSS scrollbar and layout rules verified!");

    console.log("\n=== ALL INTEGRATION TESTS PASSED SUCCESSFULLY! ===");
}

runTest().catch(e => {
    console.error("Test failed with error:", e);
    process.exit(1);
});
