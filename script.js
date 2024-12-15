// // Global array to store the combined data
// let combinedData = [];

// // Handle file selection
// document.getElementById("inputExcelFiles").addEventListener("change", function (e) {
//   const files = e.target.files;
//   if (files.length === 0) return;

//   // Clear existing data
//   combinedData = [];

//   // Clear the displayed table
//   document.getElementById("output").innerHTML = "";

//   // Read each selected file
//   for (let i = 0; i < files.length; i++) {
//     const file = files[i];
//     const reader = new FileReader();

//     reader.onload = function (event) {
//       const data = event.target.result;

//       // Parse the Excel file
//       const workbook = XLSX.read(data, { type: "binary" });

//       // Loop through all sheets
//       workbook.SheetNames.forEach((sheetName) => {
//         const sheet = workbook.Sheets[sheetName];
//         const jsonData = XLSX.utils.sheet_to_json(sheet, {
//           header: 1,
//           defval: "",
//         }); // Convert sheet to JSON

//         // Clean the data by removing rows with missing IDs in column A
//         const cleanData = cleanExtraneousData(jsonData);

//         // Combine data (you can customize how you combine the data)
//         combinedData.push({ sheetName, data: cleanData, sheet: sheet });

//         // Optionally display the sheet data on the screen (if needed)
//         displaySheet(sheetName, cleanData);
//       });
//     };

//     // Read the file as binary string
//     reader.readAsBinaryString(file);
//   }
// });

// // Clean up extraneous data: Remove rows where the ID is missing (in column A)
// function cleanExtraneousData(data) {
//   return data.filter(row => row[0] !== "" && row[0] !== undefined); // Filter out rows without ID
// }

// // Display the data in a table format on the page (optional)
// function displaySheet(sheetName, data) {
//   const outputDiv = document.getElementById("output");
//   const table = document.createElement("table");
//   table.setAttribute("border", "1");
//   table.innerHTML = `<tr><th colspan="${data[0].length}">${sheetName}</th></tr>`;

//   // Create table rows
//   data.forEach((row) => {
//     const tr = document.createElement("tr");
//     row.forEach((cell) => {
//       const td = document.createElement("td");
//       td.innerText = cell;
//       tr.appendChild(td);
//     });
//     table.appendChild(tr);
//   });

//   outputDiv.appendChild(table);
// }

// // Handle the download of combined data
// document.getElementById("downloadBtn").addEventListener("click", function () {
//   // Combine all data into one sheet
//   const combinedSheetData = [];

//   // Add header rows to the combined sheet (use the header of the first sheet)
//   if (combinedData.length > 0) {
//     const firstSheet = combinedData[0].data;
//     const headerRow = firstSheet[0]; // First header row (from the first uploaded sheet)

//     // Add header once
//     combinedSheetData.push(headerRow);
//   }

//   // Loop through each sheet's data and add to the combined data
//   combinedData.forEach((sheet) => {
//     sheet.data.slice(1).forEach((row) => {  // Skip the header row (first row)
//       combinedSheetData.push(row); // Add the row to the combined data
//     });
//   });

//   // Create a new worksheet with the combined data
//   const newWorksheet = XLSX.utils.aoa_to_sheet(combinedSheetData);

//   // Create a new workbook
//   const newWorkbook = XLSX.utils.book_new();
//   XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "CombinedSheet");

//   // Download the new Excel file with the combined data
//   XLSX.writeFile(newWorkbook, "combined_excel.xlsx");
// });

// Global array to store the combined data
let combinedData = [];

// Handle file selection
document
  .getElementById("inputExcelFiles")
  .addEventListener("change", function (e) {
    const files = e.target.files;
    if (files.length === 0) return;

    // Clear existing data
    combinedData = [];

    // Clear the displayed table
    document.getElementById("output").innerHTML = "";

    // Read each selected file
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const reader = new FileReader();

      reader.onload = function (event) {
        const data = event.target.result;

        // Parse the Excel file
        const workbook = XLSX.read(data, { type: "binary" });

        // Loop through all sheets
        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
          }); // Convert sheet to JSON

          // Clean the data by removing rows with missing IDs in column A
          const cleanData = cleanExtraneousData(jsonData);

          // Combine data (we only need the data, not the header)
          if (combinedData.length === 0) {
            // Add the first header to the combined data
            combinedData.push(cleanData[0]);
          }

          // Add the rows from this sheet (excluding the header)
          combinedData.push(...cleanData.slice(1)); // Skip the header row
        });

        // After processing all sheets, display the data in the browser
        displayCombinedData();
      };

      // Read the file as binary string
      reader.readAsBinaryString(file);
    }
  });

// Clean up extraneous data: Remove rows where the ID is missing (in column A)
function cleanExtraneousData(data) {
  return data.filter((row) => row[0] !== "" && row[0] !== undefined); 
}

// Display the combined data in a table
function displayCombinedData() {
  const outputDiv = document.getElementById("output");
  const table = document.createElement("table");
  table.setAttribute("border", "1");

  // Create the header row in the table
  const headerRow = combinedData[0];
  const headerRowElement = document.createElement("tr");
  headerRow.forEach((cell) => {
    const th = document.createElement("th");
    th.innerText = cell;
    headerRowElement.appendChild(th);
  });
  table.appendChild(headerRowElement);

  // Add the rest of the rows
  for (let i = 1; i < combinedData.length; i++) {
    const row = combinedData[i];
    const rowElement = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.innerText = cell;
      rowElement.appendChild(td);
    });
    table.appendChild(rowElement);
  }

  // Append the table to the output div
  outputDiv.appendChild(table);
}

// Handle the download of combined data
document.getElementById("downloadBtn").addEventListener("click", function () {
  // Create a new worksheet with the combined data
  const newWorksheet = XLSX.utils.aoa_to_sheet(combinedData);

  // Create a new workbook
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "CombinedSheet");

  // Download the new Excel file with the combined data
  XLSX.writeFile(newWorkbook, "combined_excel.xlsx");
});
