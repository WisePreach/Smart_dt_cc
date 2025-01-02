/****************************************************************************
 * Global Variables & Configuration
 ****************************************************************************/

// Exact column headers in your Excel file (Row 1):
const columnHeaders = [
  "L1_Disposition__c",
  "L2_Disposition__c",
  "L3_Disposition__c",
  "Restaurant_Disposition__c",
  "DE_Disposition__c",
  "Customer_Disposition__c",
  "Disposition_Card__c"
];

// Corresponding <select> element IDs in the HTML:
const dropdownIds = ["col1", "col2", "col3", "col4", "col5", "col6"];

// Excel data loaded from the file
let excelData = [];

/****************************************************************************
 * Load the Excel File Automatically on Page Load
 ****************************************************************************/
window.addEventListener("DOMContentLoaded", () => {
  console.log("DOM content loaded - starting to load Excel file...");
  loadExcelFile("DynamicNBA.xlsx");
  
  // Set up Reset button once the DOM is ready
  setupResetButton();
});

/****************************************************************************
 * 1. Load & Parse the Excel File (via XHR)
 ****************************************************************************/
function loadExcelFile(filePath) {
  const xhr = new XMLHttpRequest();
  xhr.open("GET", filePath, true);
  xhr.responseType = "arraybuffer";

  xhr.onload = function () {
    if (xhr.status >= 200 && xhr.status < 300) {
      console.log("Excel file loaded successfully. Parsing workbook...");

      const workbook = XLSX.read(xhr.response, { type: "array" });
      // Load the first sheet by default:
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Convert to JSON (each row => an object with keys = headers)
      // defval: "" ensures empty cells become empty string instead of undefined
      excelData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      console.log("Data parsed. Total rows:", excelData.length);

      // Now that data is loaded, initialize the dropdowns and listeners
      setupDropdownListeners();
      initializeDropdowns();
    } else {
      console.error("Failed to load Excel file:", xhr.statusText);
      document.getElementById("result").textContent =
        "Error: Unable to load DynamicNBA.xlsx.";
    }
  };

  xhr.onerror = function () {
    console.error("Error loading the Excel file.");
    document.getElementById("result").textContent =
      "Error: Unable to load DynamicNBA.xlsx.";
  };

  xhr.send();
}

/****************************************************************************
 * 2. Attach "change" event listeners to all dropdowns
 ****************************************************************************/
function setupDropdownListeners() {
  dropdownIds.forEach((dropdownId, index) => {
    const dropdownElem = document.getElementById(dropdownId);
    dropdownElem.addEventListener("change", () => handleSelection(index));
  });
  console.log("Dropdown listeners attached.");
}

/****************************************************************************
 * 3. Initialize the first dropdown
 ****************************************************************************/
function initializeDropdowns() {
  console.log("Initializing dropdowns...");

  // 1) Clear all dropdowns
  dropdownIds.forEach((id) => {
    document.getElementById(id).innerHTML = "<option value=''>--Select--</option>";
  });

  // 2) If we have data, populate the first dropdown
  if (excelData.length > 0) {
    const colName = columnHeaders[0]; // L1_Disposition__c
    const uniqueValues = getUniqueValues(excelData, colName);
    populateDropdown(dropdownIds[0], uniqueValues);
    console.log("First dropdown populated with", uniqueValues.length, "unique values.");
  } else {
    console.warn("No data found to populate the dropdowns.");
  }

  // 3) Clear the output text
  document.getElementById("result").textContent = "No result yet";
}

/****************************************************************************
 * 4. Populate a given dropdown with an array of values
 ****************************************************************************/
function populateDropdown(dropdownId, values) {
  const dropdown = document.getElementById(dropdownId);
  dropdown.innerHTML = "<option value=''>--Select--</option>";

  values.forEach((val) => {
    const option = document.createElement("option");
    option.value = val;
    option.textContent = val;
    dropdown.appendChild(option);
  });
}

/****************************************************************************
 * 5. Return unique, non-empty, sorted values from a column
 ****************************************************************************/
function getUniqueValues(dataArray, colName) {
  const valSet = new Set();
  dataArray.forEach((row) => {
    const cellVal = row[colName] ? row[colName].toString().trim() : "";
    if (cellVal !== "") {
      valSet.add(cellVal);
    }
  });
  return [...valSet].sort();
}

/****************************************************************************
 * 6. Handle each dropdown selection in a hierarchical manner
 ****************************************************************************/
function handleSelection(changedIndex) {
  console.log(`Dropdown ${changedIndex + 1} changed. Applying filters...`);

  // 1) Gather the current selections for all 6 columns
  const filters = dropdownIds.map((ddId) => {
    const val = document.getElementById(ddId).value;
    return val || ""; // blank if not selected
  });

  // 2) Perform strict filter to see if exactly 1 row matches
  const strictlyMatchedRows = strictFilter(excelData, filters);
  console.log("Strict filter found", strictlyMatchedRows.length, "matching rows.");

  // If exactly 1 row matches, show it
  if (strictlyMatchedRows.length === 1) {
    const lineBreakContent = strictlyMatchedRows[0][columnHeaders[6]] || "";
    document.getElementById("result").textContent = lineBreakContent;
  } else {
    document.getElementById("result").textContent = "";
  }

  // 3) Clear subsequent dropdowns (anything after changedIndex)
  clearSubsequentDropdowns(changedIndex + 1);

  // 4) Populate the next dropdown(s) if needed via relaxed filter
  let relaxedData = relaxedFilter(excelData, filters);

  for (let i = changedIndex + 1; i < dropdownIds.length; i++) {
    const nextColName = columnHeaders[i];
    const uniqueVals = getUniqueValues(relaxedData, nextColName);
    populateDropdown(dropdownIds[i], uniqueVals);

    // If a value was already selected in that dropdown, refine relaxedData
    const chosenVal = document.getElementById(dropdownIds[i]).value;
    if (chosenVal) {
      relaxedData = relaxedData.filter(
        (row) => (row[nextColName] || "").toString().trim() === chosenVal
      );
    }
  }
}

/****************************************************************************
 * 7. strictFilter
 *    For each selected column:
 *      - If userVal is blank => that column in the row must be blank
 *      - If userVal is not blank => that column in the row must match
 ****************************************************************************/
function strictFilter(dataArray, filters) {
  return dataArray.filter((row) => {
    for (let i = 0; i < filters.length; i++) {
      const userVal = filters[i];
      const cellVal = (row[columnHeaders[i]] || "").toString().trim();

      if (userVal === "") {
        // Column must be blank
        if (cellVal !== "") return false;
      } else {
        // Column must match
        if (cellVal !== userVal) return false;
      }
    }
    return true;
  });
}

/****************************************************************************
 * 8. relaxedFilter
 *    For each selected column:
 *      - If userVal is not blank => row must match that value
 *      - If userVal is blank => ignore
 ****************************************************************************/
function relaxedFilter(dataArray, filters) {
  return dataArray.filter((row) => {
    for (let i = 0; i < filters.length; i++) {
      const userVal = filters[i];
      if (userVal) {
        const cellVal = (row[columnHeaders[i]] || "").toString().trim();
        if (cellVal !== userVal) {
          return false;
        }
      }
    }
    return true;
  });
}

/****************************************************************************
 * 9. Clear subsequent dropdowns (from "startIndex" onward)
 ****************************************************************************/
function clearSubsequentDropdowns(startIndex) {
  console.log("Clearing dropdowns from index", startIndex, "to end.");
  for (let i = startIndex; i < dropdownIds.length; i++) {
    document.getElementById(dropdownIds[i]).innerHTML =
      "<option value=''>--Select--</option>";
  }
}

/****************************************************************************
 * 10. RESET Button Functionality
 ****************************************************************************/
function setupResetButton() {
  const btn = document.getElementById("resetBtn");
  if (!btn) {
    console.error("Reset button not found in HTML (id='resetBtn').");
    return;
  }
  btn.addEventListener("click", () => {
    console.log("Reset button clicked. Re-initializing dropdowns...");
    initializeDropdowns();
  });
}
