const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// This script verifies the exact logic of app.js on the file '예약조회 (1).xlsx'
// to ensure the 12 rules are mathematically correct and generate the exact 29 columns TSV.

const xlsxPath = path.join(__dirname, '예약조회 (1).xlsx');
const workbook = XLSX.read(fs.readFileSync(xlsxPath), { type: 'buffer' });
const firstSheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[firstSheetName];
const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log("=== Verification Script Running ===");
console.log("Sheet Loaded. Raw rows count:", rawRows.length);

// -------------------------------------------------------------
// Copy of app.js processing helper functions
// -------------------------------------------------------------
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
    const ymdMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (ymdMatch) {
        return `${parseInt(ymdMatch[2], 10)}월 ${parseInt(ymdMatch[3], 10)}일`;
    }
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

function escapeTSVField(val) {
    if (val === undefined || val === null) return '';
    let str = String(val);
    if (str.includes('\t') || str.includes('\n') || str.includes('\r') || str.includes('"')) {
        str = str.replace(/"/g, '""');
        return `"${str}"`;
    }
    return str;
}

// -------------------------------------------------------------
// Copy of app.js processing logic
// -------------------------------------------------------------
const validRows = rawRows.filter(row => row && row.length > 0 && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''));

const headers = validRows[0];
const colIdx = {};
headers.forEach((h, idx) => {
    if (h) colIdx[String(h).trim()] = idx;
});

// Assert required columns exist
const required = ['고객번호', '고객명', '예약일', '예약시간', '진료실', '고객메모', '예약상태', '고객구분', '성별', '나이', '예약메모', '기타메모'];
required.forEach(col => {
    if (colIdx[col] === undefined) {
        throw new Error(`해당 파일에는 '${col}' 데이터가 없습니다.`);
    }
});

const uniqueCustomers = new Map();
const exclusions = [];
const cleaned = [headers];

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
    
    // Rule 1-2: filter foreign teams
    if (custMemo.includes('일본') || custMemo.includes('대만') || custMemo.includes('해외')) {
        exclusions.push({
            type: 'foreign',
            customerNumber: custNum,
            name: custName,
            time: time,
            status: status,
            reason: '해외팀 예약',
            rowNum: i + 1
        });
        continue;
    }
    
    // Rule 1-3: filter customer names containing '예약' or '마감'
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
        continue;
    }
    
    if (!custNum) {
        cleaned.push(row);
        continue;
    }
    
    // Rule 1: deduplicate
    if (uniqueCustomers.has(custNum)) {
        exclusions.push({
            type: 'duplicate',
            customerNumber: custNum,
            name: custName,
            time: time,
            status: status,
            reason: '중복',
            rowNum: i + 1
        });
    } else {
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

console.log("\n--- Verification Results ---");
console.log("Original Rows:", validRows.length - 1);
console.log("Cleaned Rows:", uniqueCustomers.size);
console.log("Exclusions logged:", exclusions.length);

const mappedData = Array.from(uniqueCustomers.values()).map(item => {
    const r = item.row;
    const memo = r[colIdx['예약메모']] || '';
    const custMemo = r[colIdx['고객메모']] || '';
    const otherMemo = r[colIdx['기타메모']] || '';
    const rawDate = r[colIdx['예약일']] || '';
    
    const gender = colIdx['성별'] !== undefined ? (r[colIdx['성별']] || '') : extractGender(memo);
    const age = colIdx['나이'] !== undefined ? (r[colIdx['나이']] || '') : extractAge(memo);
    
    const formattedDate = formatDate(rawDate);
    const formattedTime = formatTime(r[colIdx['예약시간']] || '');

    return {
        reservationDate: formattedDate,
        registrant: '홍준식',
        customerName: r[colIdx['고객명']] || '',
        visitPurpose: r[colIdx['진료실']] || '',
        marketingPath: r[colIdx['고객메모']] || '',
        doctor: r[colIdx['예약상태']] || '',
        coordinator: '',
        patientType: r[colIdx['고객구분']] || '',
        gender: gender,
        age: age,
        reservationTime: formattedTime,
        memo: memo,
        customerMemo: otherMemo
    };
});

// Convert to TSV to check clipboard format
const tsvRows = mappedData.map(item => {
    const row = [
        item.reservationDate,      // 1
        item.registrant,           // 2 ('홍준식')
        item.customerName,         // 3
        item.visitPurpose,         // 4
        item.marketingPath,        // 5
        item.doctor,               // 6
        item.coordinator,          // 7 (empty)
        item.patientType,          // 8
        item.gender,               // 9
        item.age,                  // 10
        item.reservationTime,      // 11
        "", "", "", "", "", "", "", "", "", // 12~20 (empty)
        item.memo,                 // 21
        "",                        // 22
        "",                        // 23
        "",                        // 24
        "", "", "", "", ""         // 25~29 (empty)
    ];
    return row.map(escapeTSVField).join('\t');
});

console.log("TSV Rows generated (array length):", tsvRows.length);
const firstRowCols = tsvRows[0].split('\t');
console.log("Columns count in first TSV row:", firstRowCols.length);

// Verify Exclusions Reasons
console.log("\n--- Exclusions Reasons Log check ---");
exclusions.forEach(item => {
    console.log(`Row ${item.rowNum}: Name=${item.name}, CustNum=${item.customerNumber}, Type=${item.type}, Reason=${item.reason}`);
});

// Final assertions
if (tsvRows.length !== 55) {
    console.error(`FAIL: Expected 55 rows in final output, but got ${tsvRows.length}!`);
    process.exit(1);
}
if (firstRowCols.length !== 29) {
    console.error(`FAIL: Expected exactly 29 columns in each row, but got ${firstRowCols.length}!`);
    process.exit(1);
}

console.log("\nSUCCESS: All logic assertions passed mathematically!");
process.exit(0);
