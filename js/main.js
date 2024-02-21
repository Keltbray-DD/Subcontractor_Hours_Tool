let records = [];
let contractNumber = '';
let editIndex = -1;

function addRecord() {
    
    const company = document.getElementById('company').value;
    const name = document.getElementById('name').value;
    const hours = document.getElementById('hours').value;
    const type = document.getElementById('type').value;
    const wc = document.getElementById('wc').value;

    records.push({ company, name, hours, type, wc });
    clearForm();
    displayRecords();
}

function editRecord(index) {
    editIndex = index;
    const record = records[index];
    document.getElementById('company').value = record.company;
    document.getElementById('name').value = record.name;
    document.getElementById('hours').value = record.hours;
    document.getElementById('type').value = record.type;
    document.getElementById('wc').value = record.wc;
    
    // Apply a CSS class to the row being edited to make it appear greyed out
    const recordsBody = document.getElementById('recordsBody');
    recordsBody.childNodes[index].classList.add('editing');
    
    // Disable the Add Record button while editing
    document.getElementById('addRecordBtn').disabled = true;
    // Show the Save Changes button
    document.getElementById('saveChangesBtn').style.display = 'block';
}

function saveChanges() {
    const company = document.getElementById('company').value;
    const name = document.getElementById('name').value;
    const hours = document.getElementById('hours').value;
    const type = document.getElementById('type').value;
    const wc = document.getElementById('wc').value;

    records[editIndex] = { company, name, hours, type, wc };

    clearForm();
    displayRecords();

    // Reset editIndex
    editIndex = -1;
    // Enable the Add Record button after saving changes
    document.getElementById('addRecordBtn').disabled = false;
    // Hide the Save Changes button
    document.getElementById('saveChangesBtn').style.display = 'none';
}

function deleteRecord(index) {
    records.splice(index, 1);
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
    records.forEach((record, index) => {
      const row = document.createElement('tr');
      for (const key in record) {
        const cell = document.createElement('td');
        cell.textContent = record[key];
        row.appendChild(cell);
      }
      const actionsCell = document.createElement('td');
      const deleteBtn = document.createElement('button');
      deleteBtn.innerHTML = '<i class="material-icons">delete</i>'; // Material Icons delete icon
      deleteBtn.className = 'delete-btn';
      deleteBtn.onclick = () => deleteRecord(index);
      actionsCell.appendChild(deleteBtn);
      const editBtn = document.createElement('button');
      editBtn.innerHTML = '<i class="material-icons">edit</i>'; // Material Icons edit icon
      editBtn.className = 'edit-btn';
      editBtn.onclick = () => editRecord(index);
      actionsCell.appendChild(editBtn);
      row.appendChild(actionsCell);
      recordsBody.appendChild(row);
    });
}



function exportToExcel() {
  // Create a new workbook
  XlsxPopulate.fromBlankAsync()
    .then(workbook => {
      const sheet = workbook.sheet(0);
      
      // Set the contract number label in cell A1
      sheet.cell("A1").value("Contract Number:");

      // Set the contract number in cell B1
      sheet.cell("B1").value(contractNumber);

      // Set the contract name in cell C2
      sheet.cell("C2").value("Contract Name");

      // Add the records table starting from cell A4
      const recordsData = [
        ["Company", "Operative Name", "Hours Worked", "Type", "W.C"]
      ].concat(records.map(record => [record.company, record.name, record.hours, record.type, record.wc]));
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