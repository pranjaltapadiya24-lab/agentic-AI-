import { parseAndValidateDoctors } from './parser.js';

export function runDiagnostics() {
  const tests = [];
  
  // Helper to register test results
  const assert = (name, condition, details = "") => {
    tests.push({ name, passed: !!condition, details });
  };

  // Test Case 1: Valid Doctor Data with Default Weekends
  try {
    const mockData = [
      {
        "Doctor_Name": "Dr. Kapoor",
        "Department_Mappings": "Cardiology, Emergency",
        "Min_Hours": "40",
        "Max_Hours": "60",
        "Absent_Dates": "",
        "Min_Days_Per_Dept": "2",
        "Max_Days_Per_Dept": "5"
      }
    ];
    
    // Period: Monday 2026-08-10 to Sunday 2026-08-16 (Saturdays & Sundays are 15th & 16th)
    const { doctors, errors, warnings } = parseAndValidateDoctors(mockData, "2026-08-10", "2026-08-16");
    
    assert(
      "Valid Data Parsing & Weekend Defaults",
      doctors.length === 1 && 
      doctors[0].Doctor_Name === "Dr. Kapoor" &&
      doctors[0].Absent_Dates.includes("2026-08-15") &&
      doctors[0].Absent_Dates.includes("2026-08-16") &&
      doctors[0].Absent_Dates.length === 2 &&
      errors.length === 0,
      `Expected 1 doctor, 2 default absent days (weekends). Got: ${doctors.length} doctors, ${doctors[0]?.Absent_Dates.length} absent dates, ${errors.length} errors.`
    );
  } catch (err) {
    assert("Valid Data Parsing & Weekend Defaults", false, err.message);
  }

  // Test Case 2: Specific Absent Dates (No Weekend Defaults override)
  try {
    const mockData = [
      {
        "Doctor_Name": "Dr. Kapoor",
        "Department_Mappings": "Cardiology, Emergency",
        "Min_Hours": "40",
        "Max_Hours": "60",
        "Absent_Dates": "2026-08-12, 2026-08-13",
        "Min_Days_Per_Dept": "2",
        "Max_Days_Per_Dept": "5"
      }
    ];
    
    const { doctors, errors, warnings } = parseAndValidateDoctors(mockData, "2026-08-10", "2026-08-16");
    
    assert(
      "Custom Absent Dates Override Weekends",
      doctors.length === 1 &&
      doctors[0].Absent_Dates.includes("2026-08-12") &&
      doctors[0].Absent_Dates.includes("2026-08-13") &&
      !doctors[0].Absent_Dates.includes("2026-08-15") && // should not contain Sat/Sun
      doctors[0].Absent_Dates.length === 2 &&
      errors.length === 0,
      `Expected custom dates 12th/13th and no weekends. Got: ${JSON.stringify(doctors[0]?.Absent_Dates)}`
    );
  } catch (err) {
    assert("Custom Absent Dates Override Weekends", false, err.message);
  }

  // Test Case 3: Validation Error (Min_Hours > Max_Hours)
  try {
    const mockData = [
      {
        "Doctor_Name": "Dr. BadHours",
        "Department_Mappings": "Emergency",
        "Min_Hours": "60",
        "Max_Hours": "40",
        "Absent_Dates": "",
        "Min_Days_Per_Dept": "1",
        "Max_Days_Per_Dept": "3"
      }
    ];
    
    const { doctors, errors } = parseAndValidateDoctors(mockData, "2026-08-10", "2026-08-16");
    
    assert(
      "Detect Min_Hours > Max_Hours",
      doctors.length === 0 && errors.length === 1 && errors[0].message.includes("cannot be greater than Max_Hours"),
      `Expected 0 doctors and 1 error. Got: ${doctors.length} doctors, ${errors.length} errors: ${errors[0]?.message}`
    );
  } catch (err) {
    assert("Detect Min_Hours > Max_Hours", false, err.message);
  }

  // Test Case 4: Duplicate Doctor Name Detection
  try {
    const mockData = [
      {
        "Doctor_Name": "Dr. Duplicate",
        "Department_Mappings": "Emergency",
        "Min_Hours": "20",
        "Max_Hours": "40",
        "Absent_Dates": "",
        "Min_Days_Per_Dept": "1",
        "Max_Days_Per_Dept": "3"
      },
      {
        "Doctor_Name": "Dr. Duplicate",
        "Department_Mappings": "Cardiology",
        "Min_Hours": "20",
        "Max_Hours": "40",
        "Absent_Dates": "",
        "Min_Days_Per_Dept": "1",
        "Max_Days_Per_Dept": "3"
      }
    ];
    
    const { doctors, errors } = parseAndValidateDoctors(mockData, "2026-08-10", "2026-08-16");
    
    assert(
      "Detect Duplicate Doctor Names",
      doctors.length === 1 && errors.length === 1 && errors[0].message.includes("Duplicate doctor name found"),
      `Expected 1 doctor, 1 duplicate name error. Got: ${doctors.length} doctors, ${errors.length} errors: ${errors[0]?.message}`
    );
  } catch (err) {
    assert("Detect Duplicate Doctor Names", false, err.message);
  }

  // Test Case 5: Missing Required Columns
  try {
    const mockData = [
      {
        "Doctor_Name": "Dr. MissingCols",
        "Min_Hours": "20"
      }
    ];
    
    const { doctors, errors } = parseAndValidateDoctors(mockData, "2026-08-10", "2026-08-16");
    
    assert(
      "Detect Missing Required Columns",
      doctors.length === 0 && errors.length === 1 && errors[0].message.includes("Missing required columns"),
      `Expected 0 doctors, 1 critical column error. Got: ${doctors.length} doctors, ${errors.length} errors: ${errors[0]?.message}`
    );
  } catch (err) {
    assert("Detect Missing Required Columns", false, err.message);
  }

  return tests;
}
