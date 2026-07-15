import { parseAndValidateDoctors } from './parser.js';
import { runDiagnostics } from './tests.js';

// Global state
let currentRawData = null;
let currentFileName = "";

// Initialize Lucide Icons on load
document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    window.lucide.createIcons();
  }
  setupEventListeners();
});

function setupEventListeners() {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const startDateInput = document.getElementById('start-date');
  const endDateInput = document.getElementById('end-date');
  const shiftHoursInput = document.getElementById('shift-hours');
  const weekendAbsentCheck = document.getElementById('weekend-absent');
  const btnReset = document.getElementById('btn-reset');
  const btnDownloadTemplate = document.getElementById('btn-download-template');
  const btnRunTests = document.getElementById('btn-run-tests');
  const diagnosticResults = document.getElementById('diagnostic-results');

  // Diagnostic Test Runner
  btnRunTests.addEventListener('click', () => {
    diagnosticResults.style.display = 'block';
    diagnosticResults.innerHTML = '<div style="color: var(--color-primary); font-size: 0.75rem;">Running diagnostic suite...</div>';
    
    setTimeout(() => {
      try {
        const testSuite = runDiagnostics();
        let passCount = 0;
        let failCount = 0;
        
        let htmlOutput = '<div style="font-weight: 600; margin-bottom: 0.5rem; text-transform: uppercase; font-size: 0.6875rem; letter-spacing: 0.05em; color: var(--color-text-muted);">Diagnostic Suite Run:</div>';
        
        testSuite.forEach(t => {
          if (t.passed) {
            passCount++;
            htmlOutput += `<div style="color: var(--color-success); margin-bottom: 0.25rem; font-weight: 500;">✔ ${t.name}</div>`;
          } else {
            failCount++;
            htmlOutput += `<div style="color: var(--color-error); margin-bottom: 0.25rem; font-weight: 500;">✘ ${t.name}<br><small style="color: var(--color-text-muted); display: block; margin-left: 1.25rem; font-size: 0.75rem;">${t.details}</small></div>`;
          }
        });
        
        htmlOutput += `<div style="margin-top: 0.5rem; border-top: 1px solid var(--border-card); padding-top: 0.5rem; font-weight: 600; color: ${failCount > 0 ? 'var(--color-error)' : 'var(--color-success)'};">Status: ${passCount}/${testSuite.length} Passed</div>`;
        
        diagnosticResults.innerHTML = htmlOutput;
      } catch (err) {
        diagnosticResults.innerHTML = `<div style="color: var(--color-error);">Runner Error: ${err.message}</div>`;
      }
    }, 300);
  });


  // Drag and Drop
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
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileSelection(files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    const files = e.target.files;
    if (files.length > 0) {
      handleFileSelection(files[0]);
    }
  });

  // Re-run validation on setting changes
  const handleSettingChange = () => {
    if (currentRawData) {
      processParsedData(currentRawData);
    }
  };

  startDateInput.addEventListener('change', handleSettingChange);
  endDateInput.addEventListener('change', handleSettingChange);
  shiftHoursInput.addEventListener('change', handleSettingChange);
  weekendAbsentCheck.addEventListener('change', handleSettingChange);

  // Clear data
  btnReset.addEventListener('click', resetApp);

  // Download template
  btnDownloadTemplate.addEventListener('click', downloadTemplateExcel);
}

function handleFileSelection(file) {
  currentFileName = file.name;
  const reader = new FileReader();
  
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = window.XLSX.read(data, { type: 'array' });
      
      // Get first sheet
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Convert to JSON
      const json = window.XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      currentRawData = json;
      
      processParsedData(json);
    } catch (err) {
      console.error(err);
      showErrorAlert(`Failed to read file: ${err.message}`);
    }
  };
  
  reader.readAsArrayBuffer(file);
}

function processParsedData(rawData) {
  const startPeriod = document.getElementById('start-date').value;
  const endPeriod = document.getElementById('end-date').value;
  const useWeekends = document.getElementById('weekend-absent').checked;
  
  // Parse and Validate
  const result = parseAndValidateDoctors(
    rawData, 
    startPeriod, 
    endPeriod
  );
  
  // Render results
  renderValidationSummary(result);
  renderPreview(result);
  
  // Dynamic Icon refresh
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function renderValidationSummary(result) {
  const panel = document.getElementById('validation-panel');
  const container = document.getElementById('validation-messages');
  container.innerHTML = '';
  
  const hasErrors = result.errors.length > 0;
  const hasWarnings = result.warnings.length > 0;
  
  if (!hasErrors && !hasWarnings) {
    panel.style.display = 'block';
    container.innerHTML = `
      <div class="alert-box success">
        <i data-lucide="check-circle" class="alert-icon"></i>
        <div class="alert-content">
          <div class="alert-title">Spreadsheet Parsing Perfect!</div>
          <div>All ${result.doctors.length} doctors parsed successfully with zero warnings or structural errors. Ready for scheduling.</div>
        </div>
      </div>
    `;
    return;
  }
  
  panel.style.display = 'block';
  
  // Render errors first
  if (hasErrors) {
    const errorList = result.errors.map(err => 
      `<li>${err.doctor ? `<strong>${err.doctor}</strong> (Row ${err.row}): ` : `Row ${err.row}: `}${err.message}</li>`
    ).join('');
    
    container.innerHTML += `
      <div class="alert-box error">
        <i data-lucide="alert-octagon" class="alert-icon"></i>
        <div class="alert-content">
          <div class="alert-title">Critical Ingestion Errors Found (${result.errors.length})</div>
          <div>The following constraints are invalid and will block schedule generation:</div>
          <ul class="alert-list">${errorList}</ul>
        </div>
      </div>
    `;
  }
  
  // Render warnings
  if (hasWarnings) {
    const warningList = result.warnings.map(warn => 
      `<li>${warn.doctor ? `<strong>${warn.doctor}</strong>: ` : ''}${warn.message}</li>`
    ).join('');
    
    container.innerHTML += `
      <div class="alert-box warning">
        <i data-lucide="alert-triangle" class="alert-icon"></i>
        <div class="alert-content">
          <div class="alert-title">Parser Warnings (${result.warnings.length})</div>
          <div>Review these automatic defaults or logical mismatches:</div>
          <ul class="alert-list">${warningList}</ul>
        </div>
      </div>
    `;
  }
}

function renderPreview(result) {
  const emptyState = document.getElementById('empty-state');
  const previewContent = document.getElementById('preview-content');
  const tableBody = document.getElementById('roster-table-body');
  const proceedBtn = document.getElementById('btn-proceed-solver');
  
  // Stats
  const statDoctors = document.getElementById('stat-total-doctors');
  const statDepts = document.getElementById('stat-total-depts');
  const statWarnings = document.getElementById('stat-total-warnings');
  
  emptyState.style.display = 'none';
  previewContent.style.display = 'block';
  tableBody.innerHTML = '';
  
  // Update Stats
  statDoctors.textContent = result.doctors.length;
  
  const allDepts = new Set();
  result.doctors.forEach(doc => {
    doc.Department_Mappings.forEach(dept => allDepts.add(dept));
  });
  statDepts.textContent = allDepts.size;
  statWarnings.textContent = result.warnings.length;
  
  // Enable proceed button ONLY if there are no errors
  const hasErrors = result.errors.length > 0;
  proceedBtn.disabled = hasErrors || result.doctors.length === 0;
  
  // Populating Table
  result.doctors.forEach(doc => {
    const row = document.createElement('tr');
    
    // Check if this doctor had any specific warning
    const docHasWarning = result.warnings.some(w => w.doctor === doc.Doctor_Name);
    if (docHasWarning) {
      row.classList.add('row-warning');
    }
    
    // Departments badges
    const deptsHtml = doc.Department_Mappings.map(dept => 
      `<span class="badge-tag dept">${dept}</span>`
    ).join('');
    
    // Absent Dates badge formatting
    const absentHtml = doc.Absent_Dates.length > 0 
      ? doc.Absent_Dates.map(date => `<span class="badge-tag">${date}</span>`).join('')
      : '<span class="color-text-muted" style="font-size:0.75rem;">None</span>';
      
    row.innerHTML = `
      <td><strong>${doc.Doctor_Name}</strong></td>
      <td>${deptsHtml}</td>
      <td>${doc.Min_Hours} to ${doc.Max_Hours} hrs</td>
      <td>${doc.Min_Days_Per_Dept} to ${doc.Max_Days_Per_Dept} days</td>
      <td><div style="max-width: 250px; overflow-x: auto; display: flex; gap: 0.25rem;">${absentHtml}</div></td>
    `;
    
    tableBody.appendChild(row);
  });
}

function showErrorAlert(message) {
  const panel = document.getElementById('validation-panel');
  const container = document.getElementById('validation-messages');
  
  panel.style.display = 'block';
  container.innerHTML = `
    <div class="alert-box error">
      <i data-lucide="alert-octagon" class="alert-icon"></i>
      <div class="alert-content">
        <div class="alert-title">Critical File Ingestion Error</div>
        <div>${message}</div>
      </div>
    </div>
  `;
}

function resetApp() {
  currentRawData = null;
  currentFileName = "";
  
  document.getElementById('file-input').value = '';
  document.getElementById('validation-panel').style.display = 'none';
  document.getElementById('validation-messages').innerHTML = '';
  
  document.getElementById('empty-state').style.display = 'flex';
  document.getElementById('preview-content').style.display = 'none';
  document.getElementById('roster-table-body').innerHTML = '';
  document.getElementById('btn-proceed-solver').disabled = true;
  
  const diagnosticResults = document.getElementById('diagnostic-results');
  if (diagnosticResults) {
    diagnosticResults.style.display = 'none';
    diagnosticResults.innerHTML = '';
  }
}

function downloadTemplateExcel() {
  const headers = [
    "Doctor_Name",
    "Department_Mappings",
    "Min_Hours",
    "Max_Hours",
    "Absent_Dates",
    "Min_Days_Per_Dept",
    "Max_Days_Per_Dept"
  ];
  
  const sampleData = [
    {
      "Doctor_Name": "Dr. Kapoor",
      "Department_Mappings": "Cardiology, Emergency",
      "Min_Hours": 40,
      "Max_Hours": 60,
      "Absent_Dates": "2026-08-15, 2026-08-16",
      "Min_Days_Per_Dept": 2,
      "Max_Days_Per_Dept": 5
    },
    {
      "Doctor_Name": "Dr. Mehta",
      "Department_Mappings": "Emergency, Pediatrics",
      "Min_Hours": 32,
      "Max_Hours": 48,
      "Absent_Dates": "", // Will auto-default to Saturday/Sundays (2026-08-15, 2026-08-16)
      "Min_Days_Per_Dept": 1,
      "Max_Days_Per_Dept": 4
    },
    {
      "Doctor_Name": "Dr. Smith",
      "Department_Mappings": "Cardiology",
      "Min_Hours": 24,
      "Max_Hours": 40,
      "Absent_Dates": "2026-08-12",
      "Min_Days_Per_Dept": 2,
      "Max_Days_Per_Dept": 3
    }
  ];
  
  try {
    const worksheet = window.XLSX.utils.json_to_sheet(sampleData, { header: headers });
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "Doctors");
    
    window.XLSX.writeFile(workbook, "Aegis_Scheduler_Template.xlsx");
  } catch (err) {
    console.error(err);
    alert("Could not generate template. Please check console logs.");
  }
}
