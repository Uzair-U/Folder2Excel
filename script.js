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
  const items = dt.items;

  if (items && items.length > 0) {
    const entry = items[0].webkitGetAsEntry();
    if (entry && entry.isDirectory) {
      // We have a directory; read its immediate contents
      readDirectoryContents(entry).then(files => {
        fileList = files;
        if (fileList.length > 0) {
          generateBtn.disabled = false;
          statusEl.textContent = `Loaded ${fileList.length} files from folder.`;
        } else {
          statusEl.textContent = "No files found in the dropped folder.";
        }
      }).catch(err => {
        console.error(err);
        statusEl.textContent = "Error reading folder contents.";
      });
    } else {
      // Not a directory, just process as usual
      processFiles(dt.files);
    }
  }
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

/**
 * Read the immediate files in a directory (no recursion).
 * Continues calling readEntries() until no more entries are returned.
 * Only files are collected.
 */
function readDirectoryContents(dirEntry) {
  const reader = dirEntry.createReader();
  const allEntries = [];

  function readAllEntries() {
    return new Promise((resolve, reject) => {
      reader.readEntries(entries => {
        if (entries.length === 0) {
          resolve();
        } else {
          allEntries.push(...entries);
          // Keep reading until no entries
          readAllEntries().then(resolve).catch(reject);
        }
      }, reject);
    });
  }

  return readAllEntries().then(async () => {
    const files = [];
    for (const entry of allEntries) {
      if (entry.isFile) {
        const fileObj = await getFile(entry);
        files.push(fileObj.name);
      }
    }
    return files;
  });
}

// Convert a fileEntry to a File object
function getFile(fileEntry) {
  return new Promise((resolve, reject) => {
    fileEntry.file(file => resolve(file), reject);
  });
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
