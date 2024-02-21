let records = [];

function addRecord() {
  const company = document.getElementById('company').value;
  const name = document.getElementById('name').value;
  const hours = document.getElementById('hours').value;
  const type = document.getElementById('type').value;
  const wc = document.getElementById('wc').value;

  records.push([company, name, hours, type, wc]);
  clearForm();
  displayRecords();
}

function clearForm() {
  document.getElementById('company').value = '';
  document.getElementById('name').value = '';
  document.getElementById('hours').value = '';
  document.getElementById('type').value = 'SC';
  document.getElementById('wc').value = '';
}

function displayRecords() {
  const recordsBody = document.getElementById('recordsBody');
  recordsBody.innerHTML = '';
  records.forEach(record => {
    const row = document.createElement('tr');
    record.forEach(data => {
      const cell = document.createElement('td');
      cell.textContent = data;
      row.appendChild(cell);
    });
    recordsBody.appendChild(row);
  });
}

function exportToExcel() {
  contractNumber = document.getElementById('contractNumber').value;
  // Create a new workbook
  XlsxPopulate.fromBlankAsync()
    .then(workbook => {
      const sheet = workbook.sheet(0);
      
      // Set the contract number label in cell A1
      sheet.cell("A1").value("Contract Number:");

      // Set the contract number in cell B1
      sheet.cell("B1").value(contractNumber);

      // Add the records table starting from cell A4
      const recordsData = [
        ["Company", "Operative Name", "Hours Worked", "Type", "W.C"]
      ].concat(records);
      sheet.cell("A4").value(recordsData);

      // Generate Excel file
      return workbook.outputAsync();
    })
    .then(blob => {
      // Save the Excel file
      saveAs(blob, 'construction_records.xlsx');
    })
    .catch(error => {
      console.error('Error exporting to Excel:', error);
    });
}