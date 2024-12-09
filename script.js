const dropArea = document.getElementById('drop-area');
const folderInput = document.getElementById('folderInput');
const generateBtn = document.getElementById('generateBtn');
const statusEl = document.getElementById('status');

let fileList = [];

// Prevent default behaviors for drag events
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
  dropArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults (e) {
  e.preventDefault();
  e.stopPropagation();
}

// Highlight on dragenter and dragover
['dragenter', 'dragover'].forEach(eventName => {
  dropArea.addEventListener(eventName, () => dropArea.classList.add('hover'), false);
});

['dragleave', 'drop'].forEach(eventName => {
  dropArea.addEventListener(eventName, () => dropArea.classList.remove('hover'), false);
});

// Handle drop
dropArea.addEventListener('drop', handleDrop, false);

function handleDrop(e) {
  const dt = e.dataTransfer;
  const files = dt.files;
  processFiles(files);
}

folderInput.addEventListener('change', () => {
  const files = folderInput.files;
  processFiles(files);
});

function processFiles(files) {
  fileList = Array.from(files).map(file => file.webkitRelativePath || file.name);
  if (fileList.length > 0) {
    generateBtn.disabled = false;
    statusEl.textContent = `Loaded ${fileList.length} files.`;
  }
}

// Generate Excel file on button click
generateBtn.addEventListener('click', () => {
  if (!fileList || fileList.length === 0) {
    statusEl.textContent = "No files selected.";
    return;
  }

  // Create a new workbook and a worksheet
  const wb = XLSX.utils.book_new();
  wb.Props = {
    Title: "File List",
    Subject: "Folder Contents",
    Author: "Your Name",
    CreatedDate: new Date()
  };

  // Prepare data for the worksheet: one file name per row
  const data = [["File Name"]];
  fileList.forEach(name => data.push([name]));

  // Convert data to worksheet
  const ws = XLSX.utils.aoa_to_sheet(data);

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, "Files");

  // Generate XLSX file and force download
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'file_list.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  statusEl.textContent = "Excel file generated and downloaded.";
});
