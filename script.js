const dict = {
    en: {
        title: "Thai GPA Calculator", subtitle: "Calculate your grade average easily. Upload an Excel file, paste a table, or enter data manually.",
        box1: "1. Upload Excel File", browse: "Browse Files...", nofile: "No file chosen",
        excelRule: "⚠️ Format Required: Use the first sheet. Row 1 is ignored (for headers). Column A = Subject, Column B = Credits, Column C = Grade.",
        box2: "2. Or Paste Table Data", paste: "Copy from Word or Excel and paste here...", btnReviewPaste: "Review & Add Pasted Data",
        selectAll: "Select All", deselectAll: "Deselect All", addRow: "+ Add Row",
        thInclude: "Include", thSubject: "Subject", thCredits: "Credits", thGrade: "Grade", thRemove: "Remove",
        phSubject: "e.g. Math", phCredits: "Credits", phGrade: "Grade",
        btnCalc: "Calculate GPA", btnExport: "Export to Excel", targetGpa: "Target GPA:",
        modalTitle: "Review & Edit Data", modalDesc: "Check your imported data below before adding it to the main calculator.",
        btnCancel: "Cancel", btnAppend: "Add to Existing", btnReplace: "Replace Existing"
    },
    th: {
        title: "โปรแกรมคำนวณเกรดเฉลี่ย", subtitle: "คำนวณเกรดเฉลี่ยของคุณอย่างง่ายดาย อัปโหลดไฟล์ Excel, วางข้อมูลตาราง, หรือกรอกข้อมูลด้วยตนเอง",
        box1: "1. อัปโหลดไฟล์ Excel", browse: "เลือกไฟล์...", nofile: "ยังไม่ได้เลือกไฟล์",
        excelRule: "⚠️ รูปแบบที่รองรับ: ใช้แผ่นงานแรก แถวที่ 1 จะถูกข้าม (สำหรับหัวตาราง) คอลัมน์ A = วิชา, คอลัมน์ B = หน่วยกิต, คอลัมน์ C = เกรด",
        box2: "2. หรือวางข้อมูลตาราง", paste: "คัดลอกจาก Word หรือ Excel แล้ววางที่นี่...", btnReviewPaste: "ตรวจสอบและเพิ่มข้อมูล",
        selectAll: "เลือกทั้งหมด", deselectAll: "ยกเลิกการเลือก", addRow: "+ เพิ่มแถว",
        thInclude: "นับรวม", thSubject: "ชื่อวิชา", thCredits: "หน่วยกิต", thGrade: "เกรด", thRemove: "ลบ",
        phSubject: "เช่น คณิตศาสตร์", phCredits: "หน่วยกิต", phGrade: "เกรด",
        btnCalc: "คำนวณเกรด", btnExport: "ส่งออกเป็น Excel", targetGpa: "เกรดเฉลี่ย:",
        modalTitle: "ตรวจสอบและแก้ไขข้อมูล", modalDesc: "ตรวจสอบข้อมูลที่นำเข้าด้านล่างก่อนเพิ่มลงในเครื่องคิดเลขหลัก",
        btnCancel: "ยกเลิก", btnAppend: "เพิ่มต่อจากข้อมูลเดิม", btnReplace: "ลบข้อมูลเดิมและแทนที่"
    }
};

let currentLang = 'en';

window.onload = () => addRow();

function changeLanguage(lang) {
    currentLang = lang;
    
    // Grab everything that translates
    const allTranslatedElements = document.querySelectorAll('[data-i18n], [data-i18n-ph], #fileName');

    allTranslatedElements.forEach(el => {
        // 1. Remove the animation class
        el.classList.remove('animate-text');
        
        // 2. Force the browser to reset the animation
        void el.offsetWidth; 
        
        // 3. Change the text
        if (el.hasAttribute('data-i18n')) el.innerHTML = dict[lang][el.getAttribute('data-i18n')];
        if (el.hasAttribute('data-i18n-ph')) el.placeholder = dict[lang][el.getAttribute('data-i18n-ph')];
        
        // 4. Add the animation class back!
        el.classList.add('animate-text');
    });
    
    // Handle the file name specifically
    const fileInput = document.getElementById('uploadExcel');
    if (!fileInput.files.length) {
        document.getElementById('fileName').innerText = dict[lang].nofile;
    }
}

function addRow(subj = '', cred = '', grade = '', tableId = 'tableBody') {
    const tbody = document.getElementById(tableId);
    const tr = document.createElement('tr');
    
    let includeHtml = tableId === 'tableBody' ? `<td style="text-align: center;"><input type="checkbox" class="row-checkbox" checked onchange="calculate()"></td>` : ``;
    let onInputHtml = tableId === 'tableBody' ? `oninput="calculate()"` : ``;

    tr.innerHTML = `
        ${includeHtml}
        <td><input type="text" value="${subj}" placeholder="${dict[currentLang].phSubject}"></td>
        <td><input type="number" step="0.5" value="${cred}" class="credit-input" placeholder="${dict[currentLang].phCredits}" ${onInputHtml}></td>
        <td><input type="number" step="0.5" max="4" min="0" value="${grade}" class="grade-input" placeholder="${dict[currentLang].phGrade}" ${onInputHtml}></td>
        <td style="text-align: center;"><button onclick="this.closest('tr').remove(); ${tableId === 'tableBody' ? 'calculate();' : ''}" class="btn-remove">✖</button></td>
    `;
    tbody.appendChild(tr);
}

function toggleAll(state) {
    document.querySelectorAll('#tableBody .row-checkbox').forEach(cb => cb.checked = state);
    calculate();
}

function openModal(rows) {
    const tbody = document.getElementById('modalTableBody');
    tbody.innerHTML = '';
    rows.forEach(row => {
        if (row && row.length >= 1 && (row[0] !== '' || row[1] !== '' || row[2] !== '')) {
            addRow(row[0] || '', row[1] || '', row[2] || '', 'modalTableBody');
        }
    });
    document.getElementById('reviewModal').style.display = 'flex';
}

function closeModal() {
    document.getElementById('reviewModal').style.display = 'none';
}

function commitData(replaceExisting) {
    if (replaceExisting) {
        document.getElementById('tableBody').innerHTML = '';
    }
    document.querySelectorAll('#modalTableBody tr').forEach(tr => {
        const inputs = tr.querySelectorAll('input');
        addRow(inputs[0].value, inputs[1].value, inputs[2].value);
    });
    closeModal();
    calculate();
}

function handlePaste() {
    const text = document.getElementById('pasteArea').value;
    if (!text.trim()) return;
    const rows = text.split('\n').map(row => row.split('\t'));
    openModal(rows);
    document.getElementById('pasteArea').value = '';
}

document.getElementById('uploadExcel').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('fileName').innerText = file.name;
    const reader = new FileReader();
    reader.onload = (evt) => {
        const data = evt.target.result;
        const workbook = XLSX.read(data, {type: 'binary'});
        const firstSheet = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {header: 1});
        openModal(rows.slice(1)); // Pass rows skipping header
    };
    reader.readAsBinaryString(file);
    e.target.value = ''; // Reset input so same file triggers change again
});

function calculate() {
    let totalPoints = 0, totalCredits = 0;
    document.querySelectorAll('#tableBody tr').forEach(row => {
        if (!row.querySelector('.row-checkbox').checked) return;
        const c = parseFloat(row.querySelector('.credit-input').value);
        const g = parseFloat(row.querySelector('.grade-input').value);
        if (!isNaN(c) && !isNaN(g)) {
            totalCredits += c;
            totalPoints += (c * g);
        }
    });
    document.getElementById('gpaResult').innerText = totalCredits > 0 ? (totalPoints / totalCredits).toFixed(2) : "0.00";
}

function downloadExcel() {
    // 1. Removed the "Include" header
    const rows = [[dict[currentLang].thSubject, dict[currentLang].thCredits, dict[currentLang].thGrade]];
    
    document.querySelectorAll('#tableBody tr').forEach(tr => {
        // BONUS: Remove the "//" on the line below if you ONLY want to download rows that are checked!
        // if (!tr.querySelector('.row-checkbox').checked) return;

        const inputs = tr.querySelectorAll('input[type="text"], input[type="number"]');
        // 2. Removed the true/false check from the exported row
        rows.push([inputs[0].value, inputs[1].value, inputs[2].value]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Grades");
    XLSX.writeFile(wb, "My_GPA.xlsx");
}