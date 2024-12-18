let table = document.getElementById("output");
let combinedData = [];
let loadingSpinner = document.getElementById("loadingSpinner");
let appTitle = document.querySelector(".app-title");

// Handle file selection
document
  .getElementById("inputExcelFiles")
  .addEventListener("change", function (e) {
    const files = e.target.files;
    if (files.length === 0) return;
    loadingSpinner.style.display = "flex";

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

          // Add to combined data for download later
          if (combinedData.length === 0) {
            combinedData.push(cleanData[0]); // Add header once
          }
          combinedData.push(...cleanData.slice(1)); // Add the rows for download
          appTitle.innerHTML =
            "Data is ready! You can now download or review it.";
          // Display each sheet's data separately
          displaySheet(sheetName, cleanData);
        });
        updateTableClass();
      };

      // Read the file as binary string
      reader.readAsBinaryString(file);
    }
    loadingSpinner.style.display = "none";
  });

// Clean up extraneous data: Remove rows where the ID is missing (in column A)
function cleanExtraneousData(data) {
  return data.filter((row) => row[0] !== "" && row[0] !== undefined); // Filter out rows without ID
}

function displaySheet(sheetName, data) {
  const outputDiv = document.getElementById("output");
  const table = document.createElement("table");
  table.setAttribute("border", "1");

  // Add sheet name as a header for each table
  const sheetHeader = document.createElement("tr");
  const sheetNameCell = document.createElement("th");
  sheetNameCell.colSpan = data[0].length;
  sheetNameCell.innerText = `Sheet: ${sheetName}`;
  sheetHeader.appendChild(sheetNameCell);
  table.appendChild(sheetHeader);

  // Loop through each row and cell to display data
  data.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");

    // Loop through each cell in the row
    row.forEach((cell, colIndex) => {
      const td = document.createElement("td");

      // Check if the cell is a number (i.e., a date in Excel's serial format)
      if (colIndex !== 0 && typeof cell === "number") {
        // Convert the Excel serial number to a JavaScript Date object
        const excelDate = new Date((cell - 25569) * 86400 * 1000); // Adjust Excel date to JavaScript date
        const formattedDate = excelDate.toLocaleDateString(); // Format it as a readable date (e.g., mm/dd/yyyy)
        td.innerText = formattedDate;
      } else {
        td.innerText = cell;
      }

      tr.appendChild(td);
    });

    tr.addEventListener("click", function () {
      tr.classList.toggle("clicked");
    });
    

    table.appendChild(tr);
  });

  outputDiv.appendChild(table);
}

function updateTableClass() {
  const table = document.getElementById("output");
  if (combinedData.length > 0) {
    table.classList.add("notEmpty");
  } else {
    table.classList.remove("notEmpty");
  }
}

// Ensure that we have at least 2 files before allowing the download
document.getElementById("downloadBtn").addEventListener("click", function () {
  const files = document.getElementById("inputExcelFiles").files;

  // Check if at least two files were uploaded
  if (files.length < 2) {
    Swal.fire({
      text: "You must upload at least two files to download!",
      icon: "error",
    });
    return; // Stop the function if there are less than two files
  } else {
    let timerInterval;
    Swal.fire({
      title: "Downloading The Compined File!",
      html: "Please wait <b></b> milliseconds.",
      timer: 1500,
      timerProgressBar: true,
      didOpen: () => {
        Swal.showLoading();
        const timer = Swal.getPopup().querySelector("b");
        timerInterval = setInterval(() => {
          timer.textContent = `${Swal.getTimerLeft()}`;
        }, 100);
      },
      willClose: () => {
        clearInterval(timerInterval);
      },
    }).then((result) => {
      /* Read more about handling dismissals below */
      if (result.dismiss === Swal.DismissReason.timer) {
        downloadFile();
      }
    });
  }

  // Create a new array to store the combined data with date formatting
  const formattedData = [];

  // Loop through each row in the combinedData and format the dates properly
  combinedData.forEach((row, rowIndex) => {
    const formattedRow = row.map((cell, colIndex) => {
      if (typeof cell === "number" && colIndex !== 0) {
        // If it's a date and not the first column (ID)
        // Convert the Excel serial date to a JavaScript Date object
        const excelDate = new Date((cell - 25569) * 86400 * 1000); // Convert Excel serial to JS Date
        return excelDate; // Return the JavaScript Date object (Excel serial number will be converted)
      }
      return cell; // Return non-date cells as they are
    });

    formattedData.push(formattedRow);
  });

  // Create a new worksheet with the formatted data
  const newWorksheet = XLSX.utils.aoa_to_sheet(formattedData);

  // Format the date columns
  for (let rowIndex = 0; rowIndex < formattedData.length; rowIndex++) {
    for (
      let colIndex = 6;
      colIndex < formattedData[rowIndex].length;
      colIndex++
    ) {
      const cell =
        newWorksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];

      if (cell && cell.v instanceof Date) {
        cell.s = {
          numFmt: "dd/mm/yyyy",
        };
      }
    }
  }

  function downloadFile() {
    // Create a new workbook and append the new worksheet
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "CombinedSheet");

    // Download the new Excel file with the combined and formatted data
    XLSX.writeFile(newWorkbook, "combined_excel.xlsx");
  }
});

