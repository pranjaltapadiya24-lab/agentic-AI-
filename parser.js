/**
 * Hospital Shift Timetable Scheduler - Excel Ingestion & Validation Parser
 */

/**
 * Validates a YYYY-MM-DD date string.
 * @param {string} dateStr 
 * @returns {boolean}
 */
export function isValidDateStr(dateStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return false;
  const parts = dateStr.split('-');
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // 0-indexed month
  const day = parseInt(parts[2], 10);
  const date = new Date(year, month, day);
  return date.getFullYear() === year && date.getMonth() === month && date.getDate() === day;
}

/**
 * Helper to get all weekend dates (Saturdays and Sundays) in a given date range.
 * @param {string} startStr - YYYY-MM-DD
 * @param {string} endStr - YYYY-MM-DD
 * @returns {string[]} - Array of YYYY-MM-DD date strings
 */
export function getWeekendsInRange(startStr, endStr) {
  const weekends = [];
  const start = new Date(startStr);
  const end = new Date(endStr);
  
  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    return weekends;
  }

  let current = new Date(start);
  while (current <= end) {
    const dayOfWeek = current.getDay(); // 0 = Sunday, 6 = Saturday
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      const yyyy = current.getFullYear();
      const mm = String(current.getMonth() + 1).padStart(2, '0');
      const dd = String(current.getDate()).padStart(2, '0');
      weekends.push(`${yyyy}-${mm}-${dd}`);
    }
    current.setDate(current.getDate() + 1);
  }
  return weekends;
}

/**
 * Normalize and parse raw sheet data.
 * @param {Array<Object>} rawData - JSON array of rows from Excel/CSV
 * @param {string} startPeriod - YYYY-MM-DD
 * @param {string} endPeriod - YYYY-MM-DD
 * @returns {Object} - { doctors, errors, warnings }
 */
export function parseAndValidateDoctors(rawData, startPeriod, endPeriod) {
  const doctors = [];
  const errors = [];
  const warnings = [];
  const seenNames = new Set();

  if (!rawData || rawData.length === 0) {
    errors.push({ row: 0, message: "No data rows found in the uploaded file." });
    return { doctors, errors, warnings };
  }

  // Identify column headers and normalise keys
  const expectedKeys = {
    Doctor_Name: ['doctor_name', 'doctor name', 'doctorname', 'name', 'drname', 'dr_name'],
    Department_Mappings: ['department_mappings', 'department mappings', 'department_mapping', 'departments', 'dept_mappings', 'depts'],
    Min_Hours: ['min_hours', 'min hours', 'minhour', 'min_hour', 'minhours'],
    Max_Hours: ['max_hours', 'max hours', 'maxhour', 'max_hour', 'maxhours'],
    Absent_Dates: ['absent_dates', 'absent dates', 'absentdate', 'absent_date', 'absents'],
    Min_Days_Per_Dept: ['min_days_per_dept', 'min days per dept', 'min_days_dept', 'mindaysperdept'],
    Max_Days_Per_Dept: ['max_days_per_dept', 'max days per dept', 'max_days_dept', 'maxdaysperdept']
  };

  // Find column maps based on first row keys
  const firstRowKeys = Object.keys(rawData[0]);
  const columnMap = {};
  
  for (const [standardKey, aliases] of Object.entries(expectedKeys)) {
    const match = firstRowKeys.find(key => {
      const lowerKey = key.trim().toLowerCase();
      return lowerKey === standardKey.toLowerCase() || aliases.includes(lowerKey);
    });
    columnMap[standardKey] = match || standardKey; // fallback to standardKey if not found
  }

  // Validate columns existence
  const missingRequired = [];
  ['Doctor_Name', 'Department_Mappings'].forEach(reqKey => {
    if (!firstRowKeys.find(k => k.trim().toLowerCase() === columnMap[reqKey].toLowerCase() || expectedKeys[reqKey].includes(k.trim().toLowerCase()))) {
      missingRequired.push(reqKey);
    }
  });

  if (missingRequired.length > 0) {
    errors.push({
      row: 0,
      message: `Missing required columns: ${missingRequired.join(', ')}. Please check your spreadsheet headers.`
    });
    return { doctors, errors, warnings };
  }

  // Process rows
  rawData.forEach((row, index) => {
    const rowIndex = index + 2; // Excel rows are 1-indexed, first row is header
    const rawDoctorName = String(row[columnMap.Doctor_Name] || '').trim();
    
    if (!rawDoctorName) {
      errors.push({ row: rowIndex, message: "Doctor Name is empty or missing." });
      return;
    }

    if (seenNames.has(rawDoctorName.toLowerCase())) {
      errors.push({ row: rowIndex, doctor: rawDoctorName, message: `Duplicate doctor name found: "${rawDoctorName}".` });
      return;
    }
    seenNames.add(rawDoctorName.toLowerCase());

    // 1. Department Mappings
    const rawDepts = String(row[columnMap.Department_Mappings] || '').trim();
    if (!rawDepts) {
      errors.push({ row: rowIndex, doctor: rawDoctorName, message: "Department Mappings are empty or missing." });
      return;
    }
    const depts = rawDepts.split(',').map(d => d.trim()).filter(Boolean);
    if (depts.length === 0) {
      errors.push({ row: rowIndex, doctor: rawDoctorName, message: "No valid departments specified in mappings." });
      return;
    }

    // Helper for integers
    const parseInteger = (val, fieldName, defaultValue) => {
      if (val === undefined || val === null || val === '') {
        warnings.push({ row: rowIndex, doctor: rawDoctorName, message: `"${fieldName}" was empty. Defaulted to ${defaultValue}.` });
        return defaultValue;
      }
      const parsed = parseInt(val, 10);
      if (isNaN(parsed) || parsed < 0) {
        errors.push({ row: rowIndex, doctor: rawDoctorName, message: `"${fieldName}" must be a non-negative integer. Got: "${val}".` });
        return null;
      }
      return parsed;
    };

    // 2. Hours constraints
    const minHours = parseInteger(row[columnMap.Min_Hours], 'Min_Hours', 0);
    const maxHours = parseInteger(row[columnMap.Max_Hours], 'Max_Hours', 168); // 168 hours in a week max fallback
    if (minHours === null || maxHours === null) return; // skip row on error

    if (minHours > maxHours) {
      errors.push({
        row: rowIndex,
        doctor: rawDoctorName,
        message: `Min_Hours (${minHours}) cannot be greater than Max_Hours (${maxHours}).`
      });
      return;
    }

    // 3. Department Days constraints
    const minDaysDept = parseInteger(row[columnMap.Min_Days_Per_Dept], 'Min_Days_Per_Dept', 0);
    const maxDaysDept = parseInteger(row[columnMap.Max_Days_Per_Dept], 'Max_Days_Per_Dept', 31); // 31 days in a month max fallback
    if (minDaysDept === null || maxDaysDept === null) return; // skip row on error

    if (minDaysDept > maxDaysDept) {
      errors.push({
        row: rowIndex,
        doctor: rawDoctorName,
        message: `Min_Days_Per_Dept (${minDaysDept}) cannot be greater than Max_Days_Per_Dept (${maxDaysDept}).`
      });
      return;
    }

    // 4. Absent Dates & Weekend Calculation
    const rawAbsentDates = String(row[columnMap.Absent_Dates] || '').trim();
    let absentDates = [];
    
    if (rawAbsentDates) {
      const splitDates = rawAbsentDates.split(',').map(d => d.trim()).filter(Boolean);
      for (const d of splitDates) {
        if (!isValidDateStr(d)) {
          errors.push({
            row: rowIndex,
            doctor: rawDoctorName,
            message: `Invalid date format in Absent_Dates: "${d}". Use YYYY-MM-DD (e.g. 2026-08-15).`
          });
          return;
        }
        // Verify date lies in range (if date range is set)
        if (startPeriod && endPeriod) {
          const checkDate = new Date(d);
          const start = new Date(startPeriod);
          const end = new Date(endPeriod);
          if (checkDate < start || checkDate > end) {
            warnings.push({
              row: rowIndex,
              doctor: rawDoctorName,
              message: `Absent date "${d}" is outside the selected scheduling range (${startPeriod} to ${endPeriod}).`
            });
          }
        }
        absentDates.push(d);
      }
    } else {
      // Sat/Sun defaults
      if (startPeriod && endPeriod) {
        absentDates = getWeekendsInRange(startPeriod, endPeriod);
        warnings.push({
          row: rowIndex,
          doctor: rawDoctorName,
          message: `No Absent_Dates provided. Saturdays and Sundays in range default to absent (${absentDates.length} days total).`
        });
      } else {
        warnings.push({
          row: rowIndex,
          doctor: rawDoctorName,
          message: "No Absent_Dates provided. Weekends could not be auto-calculated because no scheduling range is active."
        });
      }
    }

    // Deduplicate absent dates
    absentDates = [...new Set(absentDates)].sort();

    doctors.push({
      Doctor_Name: rawDoctorName,
      Department_Mappings: depts,
      Min_Hours: minHours,
      Max_Hours: maxHours,
      Absent_Dates: absentDates,
      Min_Days_Per_Dept: minDaysDept,
      Max_Days_Per_Dept: maxDaysDept
    });
  });

  return { doctors, errors, warnings };
}
