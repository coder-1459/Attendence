/* ===== Storage Keys ===== */
const STORAGE_KEYS = {
  students: 'rollcall_students',
  attendance: 'rollcall_attendance'
};

/* ===== State ===== */
let students = [];
let attendance = {};
let currentMarks = {};

/* ===== DOM References ===== */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const fileInput = $('#fileInput');
const uploadZone = $('#uploadZone');
const browseBtn = $('#browseBtn');
const uploadMessage = $('#uploadMessage');
const studentPreview = $('#studentPreview');
const studentList = $('#studentList');
const studentCount = $('#studentCount');

const attendanceDate = $('#attendanceDate');
const attendanceMessage = $('#attendanceMessage');
const noStudentsMsg = $('#noStudentsMsg');
const attendanceList = $('#attendanceList');
const attendanceActions = $('#attendanceActions');
const saveAttendanceBtn = $('#saveAttendance');
const presentCount = $('#presentCount');
const absentCount = $('#absentCount');
const unmarkedCount = $('#unmarkedCount');

const reportMonth = $('#reportMonth');
const reportMessage = $('#reportMessage');
const noReportData = $('#noReportData');
const reportTableWrap = $('#reportTableWrap');
const reportTableBody = $('#reportTableBody');
const reportActions = $('#reportActions');
const downloadReportBtn = $('#downloadReport');

/* ===== Init ===== */
function init() {
  if (typeof XLSX === 'undefined') {
    showMessage(
      uploadMessage,
      'Excel library failed to load. Refresh the page or ensure xlsx.full.min.js is in the project folder.',
      'error'
    );
  }

  loadFromStorage();
  setDefaultDates();
  bindEvents();
  updateStudentBadge();
  renderStudentPreview();
  renderAttendance();
  generateReport();
}

function loadFromStorage() {
  try {
    const storedStudents = localStorage.getItem(STORAGE_KEYS.students);
    const storedAttendance = localStorage.getItem(STORAGE_KEYS.attendance);
    students = storedStudents ? JSON.parse(storedStudents) : [];
    attendance = storedAttendance ? JSON.parse(storedAttendance) : {};
  } catch (e) {
    students = [];
    attendance = {};
    console.error('Failed to load from localStorage:', e);
  }
}

function saveStudents() {
  localStorage.setItem(STORAGE_KEYS.students, JSON.stringify(students));
}

function saveAttendanceData() {
  localStorage.setItem(STORAGE_KEYS.attendance, JSON.stringify(attendance));
}

function setDefaultDates() {
  const today = formatDate(new Date());
  attendanceDate.value = today;
  const now = new Date();
  reportMonth.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/* ===== Tab Navigation ===== */
function switchTab(tabId) {
  $$('.tab').forEach((tab) => {
    const isActive = tab.dataset.tab === tabId;
    tab.classList.toggle('active', isActive);
    tab.setAttribute('aria-selected', isActive);
  });

  $$('.panel').forEach((panel) => {
    const isActive = panel.id === tabId;
    panel.classList.toggle('active', isActive);
    panel.hidden = !isActive;
  });

  if (tabId === 'attendance') renderAttendance();
  if (tabId === 'report') generateReport();
}

/* ===== Messages ===== */
function showMessage(el, text, type = 'info') {
  el.textContent = text;
  el.className = `message ${type}`;
  el.hidden = false;
}

function hideMessage(el) {
  el.hidden = true;
}

/* ===== Upload Students ===== */
function parseStudentFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        let workbook;
        if (file.name.endsWith('.csv')) {
          const text = e.target.result;
          workbook = XLSX.read(text, { type: 'string' });
        } else {
          const data = new Uint8Array(e.target.result);
          workbook = XLSX.read(data, { type: 'array' });
        }

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (!rows || rows.length < 2) {
          reject(new Error('File must contain a header row and at least one student.'));
          return;
        }

        const names = [];
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;
          const name = String(row[0]).trim();
          if (name) names.push(name);
        }

        if (names.length === 0) {
          reject(new Error('No valid student names found in the first column.'));
          return;
        }

        resolve(names);
      } catch (err) {
        console.error('Parse error:', err);
        reject(new Error('Could not parse file. Please upload a valid .xlsx or .csv file.'));
      }
    };

    reader.onerror = () => reject(new Error('Failed to read the file.'));

    if (file.name.endsWith('.csv')) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

async function handleFileUpload(file) {
  hideMessage(uploadMessage);

  if (typeof XLSX === 'undefined') {
    showMessage(uploadMessage, 'Excel library not loaded. Refresh the page and try again.', 'error');
    return;
  }

  if (!file) {
    showMessage(uploadMessage, 'Please select a file to upload.', 'error');
    return;
  }

  const validExt = /\.(xlsx|xls|csv)$/i;
  if (!validExt.test(file.name)) {
    showMessage(uploadMessage, 'Invalid file format. Please upload .xlsx, .xls, or .csv.', 'error');
    return;
  }

  uploadZone.classList.add('loading');

  try {
    const names = await parseStudentFile(file);
    students = names;
    saveStudents();
    updateStudentBadge();
    renderStudentPreview();
    showMessage(uploadMessage, `Successfully loaded ${students.length} student${students.length !== 1 ? 's' : ''}.`, 'success');
  } catch (err) {
    showMessage(uploadMessage, err.message, 'error');
  } finally {
    uploadZone.classList.remove('loading');
  }
}

function updateStudentBadge() {
  studentCount.textContent = students.length;
}

function renderStudentPreview() {
  if (students.length === 0) {
    studentPreview.hidden = true;
    return;
  }

  studentPreview.hidden = false;
  studentList.innerHTML = students
    .map((name) => `<li>${escapeHtml(name)}</li>`)
    .join('');
}

/* ===== Take Attendance ===== */
function loadMarksForDate(dateStr) {
  currentMarks = {};
  if (attendance[dateStr]) {
    currentMarks = { ...attendance[dateStr] };
  }
}

function renderAttendance() {
  hideMessage(attendanceMessage);
  const dateStr = attendanceDate.value;
  loadMarksForDate(dateStr);

  if (students.length === 0) {
    noStudentsMsg.hidden = false;
    attendanceList.innerHTML = '';
    attendanceActions.hidden = true;
    return;
  }

  noStudentsMsg.hidden = true;
  attendanceActions.hidden = false;

  attendanceList.innerHTML = students
    .map((name) => {
      const status = currentMarks[name] || 'unmarked';
      return `
        <div class="student-row" data-student="${escapeAttr(name)}">
          <span class="student-name">${escapeHtml(name)}</span>
          <div class="toggle-group">
            <button type="button" class="toggle-btn present ${status === 'Present' ? 'active' : ''}" data-status="Present">Present</button>
            <button type="button" class="toggle-btn absent ${status === 'Absent' ? 'active' : ''}" data-status="Absent">Absent</button>
          </div>
        </div>
      `;
    })
    .join('');

  updateAttendanceStats();
}

function handleToggleClick(e) {
  const btn = e.target.closest('.toggle-btn');
  if (!btn) return;

  const row = btn.closest('.student-row');
  const studentName = row.dataset.student;
  const status = btn.dataset.status;

  if (btn.classList.contains('active')) {
    delete currentMarks[studentName];
    btn.classList.remove('active');
  } else {
    currentMarks[studentName] = status;
    row.querySelectorAll('.toggle-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
  }

  updateAttendanceStats();
}

function updateAttendanceStats() {
  let present = 0;
  let absent = 0;
  let unmarked = 0;

  students.forEach((name) => {
    const status = currentMarks[name];
    if (status === 'Present') present++;
    else if (status === 'Absent') absent++;
    else unmarked++;
  });

  presentCount.textContent = present;
  absentCount.textContent = absent;
  unmarkedCount.textContent = unmarked;
}

function handleSaveAttendance() {
  hideMessage(attendanceMessage);

  if (students.length === 0) {
    showMessage(attendanceMessage, 'No students loaded. Please upload a student list first.', 'error');
    return;
  }

  const dateStr = attendanceDate.value;
  if (!dateStr) {
    showMessage(attendanceMessage, 'Please select a date.', 'error');
    return;
  }

  const marked = Object.keys(currentMarks);
  if (marked.length === 0) {
    showMessage(attendanceMessage, 'Please mark at least one student before saving.', 'error');
    return;
  }

  attendance[dateStr] = { ...currentMarks };
  saveAttendanceData();
  showMessage(attendanceMessage, `Attendance saved for ${dateStr} (${marked.length} student${marked.length !== 1 ? 's' : ''} marked).`, 'success');
}

/* ===== Monthly Report ===== */
function generateReport() {
  hideMessage(reportMessage);
  reportTableBody.innerHTML = '';

  const monthVal = reportMonth.value;
  if (!monthVal) {
    noReportData.hidden = false;
    reportTableWrap.hidden = true;
    reportActions.hidden = true;
    return;
  }

  const [year, month] = monthVal.split('-').map(Number);
  const prefix = `${year}-${String(month).padStart(2, '0')}`;

  const monthRecords = Object.entries(attendance).filter(([date]) => date.startsWith(prefix));

  if (students.length === 0) {
    showMessage(reportMessage, 'No students loaded. Upload a student list first.', 'info');
    noReportData.hidden = false;
    reportTableWrap.hidden = true;
    reportActions.hidden = true;
    return;
  }

  if (monthRecords.length === 0) {
    noReportData.hidden = false;
    reportTableWrap.hidden = true;
    reportActions.hidden = true;
    return;
  }

  const reportData = students.map((name) => {
    let totalDays = 0;
    let daysPresent = 0;

    monthRecords.forEach(([, dayRecord]) => {
      if (dayRecord[name]) {
        totalDays++;
        if (dayRecord[name] === 'Present') daysPresent++;
      }
    });

    const percentage = totalDays > 0 ? (daysPresent / totalDays) * 100 : 0;
    return { name, daysPresent, totalDays, percentage };
  });

  const hasData = reportData.some((r) => r.totalDays > 0);
  if (!hasData) {
    noReportData.hidden = false;
    reportTableWrap.hidden = true;
    reportActions.hidden = true;
    return;
  }

  noReportData.hidden = true;
  reportTableWrap.hidden = false;
  reportActions.hidden = false;

  reportTableBody.innerHTML = reportData
    .map((row) => {
      const pct = row.totalDays > 0 ? row.percentage.toFixed(1) : '—';
      const pctClass = row.percentage >= 80 ? 'pct-high' : row.percentage >= 60 ? 'pct-mid' : 'pct-low';
      const barWidth = row.totalDays > 0 ? row.percentage : 0;

      return `
        <tr>
          <td>${escapeHtml(row.name)}</td>
          <td>${row.daysPresent}</td>
          <td>${row.totalDays}</td>
          <td>
            <div class="pct-bar-wrap">
              <div class="pct-bar"><div class="pct-bar-fill" style="width: ${barWidth}%"></div></div>
              <span class="pct-value ${pctClass}">${pct}${row.totalDays > 0 ? '%' : ''}</span>
            </div>
          </td>
        </tr>
      `;
    })
    .join('');
}

function downloadReport() {
  hideMessage(reportMessage);

  const monthVal = reportMonth.value;
  if (!monthVal) {
    showMessage(reportMessage, 'Please select a month.', 'error');
    return;
  }

  const [year, month] = monthVal.split('-').map(Number);
  const prefix = `${year}-${String(month).padStart(2, '0')}`;
  const monthRecords = Object.entries(attendance).filter(([date]) => date.startsWith(prefix));

  if (monthRecords.length === 0 || students.length === 0) {
    showMessage(reportMessage, 'No data available to export for this month.', 'error');
    return;
  }

  const rows = [['Name', 'Days Present', 'Total Days', 'Percentage']];

  students.forEach((name) => {
    let totalDays = 0;
    let daysPresent = 0;

    monthRecords.forEach(([, dayRecord]) => {
      if (dayRecord[name]) {
        totalDays++;
        if (dayRecord[name] === 'Present') daysPresent++;
      }
    });

    const percentage = totalDays > 0 ? ((daysPresent / totalDays) * 100).toFixed(1) + '%' : 'N/A';
    rows.push([name, daysPresent, totalDays, percentage]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Attendance Report');

  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const filename = `Attendance_Report_${monthNames[month - 1]}_${year}.xlsx`;

  XLSX.writeFile(wb, filename);
  showMessage(reportMessage, `Report downloaded as ${filename}`, 'success');
}

/* ===== Utilities ===== */
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function escapeAttr(str) {
  return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ===== Event Bindings ===== */
function bindEvents() {
  $$('.tab').forEach((tab) => {
    tab.addEventListener('click', () => switchTab(tab.dataset.tab));
  });

  browseBtn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) handleFileUpload(file);
    fileInput.value = '';
  });

  uploadZone.addEventListener('click', (e) => {
    if (e.target.closest('#browseBtn')) return;
    fileInput.click();
  });

  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('dragover');
  });

  uploadZone.addEventListener('dragleave', () => {
    uploadZone.classList.remove('dragover');
  });

  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    handleFileUpload(file);
  });

  attendanceDate.addEventListener('change', renderAttendance);
  attendanceList.addEventListener('click', handleToggleClick);
  saveAttendanceBtn.addEventListener('click', handleSaveAttendance);

  reportMonth.addEventListener('change', generateReport);
  downloadReportBtn.addEventListener('click', downloadReport);

  document.addEventListener('click', (e) => {
    const gotoBtn = e.target.closest('[data-goto]');
    if (gotoBtn) switchTab(gotoBtn.dataset.goto);
  });
}

/* ===== Boot ===== */
document.addEventListener('DOMContentLoaded', init);
