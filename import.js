document.getElementById('fileInput').addEventListener('change', handleFile, false);
let currentEditingRow = null;

// Adds or ensures the override checkbox appears correctly
function ensureOverrideCheckbox() {
  if (document.getElementById('overrideLine')) return;

  const panel = document.querySelector('.form-buttons-panel');
  if (!panel) return;

  const checkboxWrap = document.createElement('div');
  checkboxWrap.style.margin = '0.5rem 0';
  checkboxWrap.style.padding = '0 0.5rem';

  const label = document.createElement('label');
  label.style.fontWeight = 'bold';
  label.style.fontSize = '13px';
  label.style.display = 'flex';
  label.style.alignItems = 'center';
  label.innerHTML = `
    <input type="checkbox" id="overrideLine" style="margin-right: 6px;" />
    Override Auto Line Number
  `;

  checkboxWrap.appendChild(label);

  const clearBtn = Array.from(panel.querySelectorAll('button')).find(btn => btn.textContent.trim().toLowerCase() === 'clear');
  const allBtns = [...panel.children];
  const hr = allBtns.find(el => el.tagName === 'HR');

  if (hr) {
    panel.insertBefore(checkboxWrap, hr);
  } else if (clearBtn) {
    panel.insertBefore(checkboxWrap, clearBtn);
  } else {
    panel.appendChild(checkboxWrap);
  }
}



function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  const isCSV = file.name.endsWith('.csv');

  const headerRowIndex = parseInt(prompt("Which row contains the column headers? (e.g., 2 or 3)")) || 1;

  reader.onload = function(e) {
    const data = e.target.result;
    if (isCSV) {
      const rows = data.split('\n').map(row => row.split(','));
      renderTable(rows.filter(r => r.length > 1).slice(headerRowIndex));
    } else {
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
      renderTable(json.filter(r => r.length > 1).slice(headerRowIndex));
    }
  };

  if (isCSV) {
    reader.readAsText(file);
  } else {
    reader.readAsBinaryString(file);
  }
}

function renderTable(data) {
  const tbody = document.querySelector('tbody');
  tbody.innerHTML = '';

  const scriptLibrary = JSON.parse(localStorage.getItem('scriptsData') || '{}');

  data.forEach((row, i) => {
    const hasContent = row.some(cell => cell && String(cell).trim() !== '');
    if (!hasContent) return;

    const tr = document.createElement('tr');
    const time1 = row[2] ? String(row[2]).padStart(4, '0') : '';
    const scriptText = row[13]?.trim() || '';

    const cells = [
      row[0] || (i + 1), // 👈 Use the saved Line# if present, fallback to row #
      row[1] || '',
      time1,
      row[3] || '',
      row[4] || '',
      row[5] || '',
      row[6] || '',
      row[7] || '',
      row[8] || '',
      row[9] || '',
      row[10] || '',
      row[11] || '',
      row[12] || '',
      scriptText,
      i + 1, // Hidden Row# for tracking
    ];

    cells.forEach((value, colIndex) => {
      const td = document.createElement('td');

      if (colIndex === 13) {
        td.classList.add('scripts-cell');

        const scrollBox = document.createElement('div');
        scrollBox.className = 'script-scrollbox';
        scrollBox.textContent = value;

        td.appendChild(scrollBox);
      } else {
        td.textContent = value;
        if (colIndex === 14) td.classList.add('hidden-col'); // Row# hidden
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  localStorage.setItem('scriptsData', JSON.stringify(scriptLibrary));
  populateScriptDropdown();
  setupRowEvents();
  setupColumnSorting();
  requestAnimationFrame(() => resizeProxyWidth());
  manageScrollVisibility();
  setTimeout(resizeProxyWidth, 0);

}






// DOMContentLoaded event

document.addEventListener('DOMContentLoaded', () => {
  const proxy = document.getElementById('horizontalScrollProxy');
  const wrapper = document.getElementById('tableScrollWrapper');

  if (proxy && wrapper) {
    proxy.addEventListener('scroll', () => {
      wrapper.scrollLeft = proxy.scrollLeft;
    });

    wrapper.addEventListener('scroll', () => {
      proxy.scrollLeft = wrapper.scrollLeft;
    });

    // Sync width
    const table = wrapper.querySelector('table');
    if (table) {
      proxy.firstElementChild.style.width = table.scrollWidth + 'px';
    }
  }
  const formFields = document.querySelector('.form-fields');
  loadSavedUnitTraining();

  const savedInjects = localStorage.getItem('mselInjectData');
  if (savedInjects) {
    try {
      const parsedInjects = JSON.parse(savedInjects);
      renderTable(parsedInjects);
      initScriptCellHeights();
    } catch (err) {
      console.warn('Failed to parse saved inject data:', err);
    }
  }

  const labels = [
    'Line #', 'Date', 'Time', 'Storyline', 'Event', 'Inject', 'Description',
    'Control Zone', 'Inject Platform', 'Method',
    'Target Audience', 'Narrative for storyline / Notes', 'T&R Event',
    'Scripts', 'Row #'
  ];

  const multilineFields = [7, 12, 13, 14];

  if (formFields && formFields.childElementCount === 0) {
    const column1 = document.createElement('div');
    const column2 = document.createElement('div');
    const column3 = document.createElement('div');
    column1.className = 'form-column';
    column2.className = 'form-column';
    column3.className = 'form-column';

    labels.forEach((labelText, i) => {
      const fieldNum = i + 1;
      if (fieldNum === 15) return;

      const label = document.createElement('label');
      label.textContent = labelText;
      let input;

      if (multilineFields.includes(fieldNum)) {
        input = document.createElement('textarea');
      } else {
        input = document.createElement('input');
        if (fieldNum === 2) input.type = 'date';
      }

      input.id = 'Field' + fieldNum;
      label.appendChild(input);

      if (i + 1 === 15) {
        label.classList.add('hidden');
        input.classList.add('hidden');
      }

        if (fieldNum === 14) {
          const dropdown = document.createElement('select');
          dropdown.id = 'scriptDropdown';
          dropdown.style.marginBottom = '0.5rem';

          const defaultOption = document.createElement('option');
          defaultOption.value = '';
          defaultOption.textContent = '-- Select a Script --';
          dropdown.appendChild(defaultOption);

          const scripts = JSON.parse(localStorage.getItem('scriptsData') || '{}');
          Object.keys(scripts).forEach(title => {
            const option = document.createElement('option');
            option.value = title;
            option.textContent = title;
            dropdown.appendChild(option);
          });

          dropdown.addEventListener('change', function () {
            const selected = this.value;
            const scriptBody = scripts[selected];
            const targetTextarea = document.getElementById('Field14');
            if (targetTextarea && scriptBody !== undefined) {
              targetTextarea.value = scriptBody;
            }
          });

          label.insertBefore(dropdown, input);  // `input` is Field13 already
          column3.appendChild(label);




        /*label.insertBefore(dropdown, input);*/
      } else if (i % 2 === 0) {
        column1.appendChild(label);
      } else {
        column2.appendChild(label);
      }
    });

    formFields.appendChild(column1);
    formFields.appendChild(column2);
    formFields.appendChild(column3);
  }

  ensureOverrideCheckbox();
  const header = document.createElement('div');
  header.id = 'formHeader';
  header.textContent = 'Inject Form';
  header.style.cursor = 'move';
  header.style.background = '#ccc';
  header.style.padding = '8px 12px';
  header.style.fontWeight = 'bold';
  header.style.position = 'relative';
  header.style.justifyContent = 'space-between';
  header.style.alignItems = 'center';

  const closeBtn = document.createElement('button');
  closeBtn.id = 'closeFormBtn';
  closeBtn.innerHTML = '&times;'; // cleaner × symbol
  closeBtn.title = 'Close Form';
  closeBtn.onclick = cancelForm;
  header.appendChild(closeBtn);
  injectForm.insertBefore(header, injectForm.firstChild);

  let offsetX = 0, offsetY = 0, isDown = false;
  injectForm.style.position = 'absolute';

  header.addEventListener('mousedown', (e) => {
    if (e.target !== header) return;
    e.preventDefault();
    isDown = true;

    const form = document.getElementById('injectForm');
    const rect = form.getBoundingClientRect();

    form.style.left = `${rect.left}px`;
    form.style.top = `${rect.top}px`;
    form.style.transform = 'none';
    form.style.position = 'absolute';
    form.style.width = `${rect.width}px`;
    form.style.maxWidth = 'none';

    offsetX = e.clientX - rect.left;
    offsetY = e.clientY - rect.top;
  });


  document.addEventListener('mouseup', () => { isDown = false; });

  document.addEventListener('mousemove', (e) => {
    if (!isDown) return;

    const form = document.getElementById('injectForm');
    const formWidth = form.offsetWidth;
    const formHeight = form.offsetHeight;
    const windowWidth = window.innerWidth;
    const windowHeight = window.innerHeight;

    // Limit dragging within viewport
    let newLeft = e.clientX - offsetX;
    let newTop = e.clientY - offsetY;

    newLeft = Math.max(0, Math.min(newLeft, windowWidth - formWidth));
    newTop = Math.max(0, Math.min(newTop, windowHeight - formHeight));

    form.style.left = `${newLeft}px`;
    form.style.top = `${newTop}px`;
  });


  /*---------------------*/
  function initScriptCellHeights() {
    document.querySelectorAll('textarea.script-cell').forEach(textarea => {
      textarea.style.height = 'auto';
      textarea.style.height = `${textarea.scrollHeight}px`;

      textarea.addEventListener('input', () => {
        textarea.style.height = 'auto';
        textarea.style.height = `${textarea.scrollHeight}px`;
      });
    });
  }

  /*-----------------*/
  populateScriptDropdown();

  const proxyInner = proxy.firstElementChild;
  const table = wrapper.querySelector('table');

  function syncScroll() {
    wrapper.scrollLeft = proxy.scrollLeft;
  }

  function mirrorScroll() {
    proxy.scrollLeft = wrapper.scrollLeft;
  }

  function updateProxySize() {
    const wrapper = document.getElementById('tableScrollWrapper');
    const proxy = document.getElementById('horizontalScrollProxy');
    const proxyInner = document.getElementById('scrollProxyInner');
    const realTable = wrapper.querySelector('table');

    const scrollNeeded = wrapper.scrollWidth > wrapper.clientWidth;
    proxy.style.display = scrollNeeded ? 'block' : 'none';
    proxyInner.style.width = table.scrollWidth + 'px';
    proxy.scrollLeft = wrapper.scrollLeft;
  }

  proxy.addEventListener('scroll', syncScroll);
  wrapper.addEventListener('scroll', mirrorScroll);
  window.addEventListener('resize', updateProxySize);

 window.addEventListener('load', () => requestAnimationFrame(updateProxySize));


  function initScriptCellHeights() {
    const scriptCells = document.querySelectorAll('textarea.script-cell');
    scriptCells.forEach(textarea => {
      textarea.style.height = 'auto'; // reset height
      textarea.style.height = textarea.scrollHeight + 'px'; // shrink to fit one line
    });
  }
 
  const scrollProxyInner = document.getElementById('scrollProxyInner');

  const resizeObserver = new ResizeObserver(() => {
    if (table && scrollProxyInner) {
      scrollProxyInner.style.width = table.scrollWidth + 'px';
    }
  });

  if (table) {
    resizeObserver.observe(table);
  }


  updateProxySize();
});

function openInjectForm() {
  document.getElementById('injectForm').style.display = 'block';
  document.getElementById('overlayBackground').style.display = 'block';
  clearForm();

  currentEditingRow = null; // ✅ Ensure it's a new inject

  // Generate override checkbox if missing
  ensureOverrideCheckbox();

  // ✅ Ensure all Field1–Field15 exist before attaching events
  const formFieldsContainer = document.querySelector('.form-fields');
  if (!formFieldsContainer.querySelector('#Field1')) {
    location.reload(); // force DOM rebuild if fields are missing
  }
}


let currentSort = { column: null, ascending: true };

function setupColumnSorting() {
  const headers = document.querySelectorAll('thead th');

  headers.forEach((th, index) => {
    if (th.classList.contains('hidden')) return;

    // Create and append a span for the sort arrow
    let arrow = document.createElement('span');
    arrow.className = 'sort-arrow';
    arrow.textContent = ' ⇅';  // Neutral arrow
    th.appendChild(arrow);

    th.style.cursor = 'pointer';
    th.addEventListener('click', () => {
      sortTableByColumn(index);
      updateSortArrows(index);
    });
  });
}

function updateSortArrows(activeIndex) {
  const headers = document.querySelectorAll('thead th');

  headers.forEach((th, i) => {
    const arrow = th.querySelector('.sort-arrow');
    if (!arrow) return;

    if (i === activeIndex) {
      arrow.textContent = currentSort.ascending ? ' ↑' : ' ↓';
    } else {
      arrow.textContent = ' ⇅'; // Reset others
    }
  });
}


function sortTableByColumn(colIndex) {
  const tbody = document.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));

  const headerText = document.querySelectorAll('thead th')[colIndex].textContent;
  const isDate = headerText.includes('Date');
  const isTime = headerText.includes('Time');
  const isNumeric = ['Line #'].includes(headerText);

  // Step 1: Split into override and normal rows
  const overrideRows = rows.filter(row => row.cells[15]?.textContent === 'true');
  const sortableRows = rows.filter(row => row.cells[15]?.textContent !== 'true');

  // Step 2: Sort only the sortable ones
  sortableRows.sort((a, b) => {
    const aText = a.children[colIndex]?.textContent.trim() || '';
    const bText = b.children[colIndex]?.textContent.trim() || '';

    if (isDate) {
      return new Date(aText) - new Date(bText);
    }

    if (isTime) {
      const pad = v => v.padStart(4, '0');
      return parseInt(pad(aText), 10) - parseInt(pad(bText), 10);
    }

    if (isNumeric || headerText.includes("Line #")) {
      const parseSmart = str => {
        const match = str.match(/^(\d+)([a-zA-Z]*)$/);
        if (!match) return [Number.MAX_SAFE_INTEGER, '']; // Push malformed to end
        return [parseInt(match[1], 10), match[2] || ''];
      };

      const [aNum, aSuffix] = parseSmart(aText);
      const [bNum, bSuffix] = parseSmart(bText);

      if (aNum !== bNum) return aNum - bNum;
      return aSuffix.localeCompare(bSuffix);
    }

    return aText.localeCompare(bText);
  });

  // Step 3: Reverse if needed
  if (currentSort.column === colIndex && currentSort.ascending) {
    sortableRows.reverse();
    currentSort.ascending = false;
  } else {
    currentSort = { column: colIndex, ascending: true };
  }

  // Step 4: Clear tbody and rebuild it
  tbody.innerHTML = '';

  // Mix override rows back in by row index
  const totalRows = [...overrideRows, ...sortableRows];
  totalRows.sort((a, b) => {
    const aOverride = a.cells[15]?.textContent === 'true';
    const bOverride = b.cells[15]?.textContent === 'true';
    return aOverride === bOverride ? 0 : aOverride ? -1 : 1;
  });

  totalRows.forEach(row => tbody.appendChild(row));

  Array.from(tbody.children).forEach((row, index) => {
  row.children[0].textContent = index + 1;       // Line#
  row.children[14].textContent = index + 1;      // Row#
});

}



function saveEditsOnly() {
  if (!currentEditingRow) {
    addNewRow(true); // treat as new entry, and close form
    return;
  }
  const cells = currentEditingRow.querySelectorAll('td');
  const overrideLine = document.getElementById('overrideLine')?.checked;

  // Update visible columns from form fields
  for (let i = 0; i < 14; i++) {
    const field = document.getElementById('Field' + (i + 1));
    if (field && cells[i]) {
      // Scripts column (index 13) uses scrollbox
      if (i === 13) {
        let scrollbox = cells[i].querySelector('.script-scrollbox');
        if (!scrollbox) {
          scrollbox = document.createElement('div');
          scrollbox.className = 'script-scrollbox';
          cells[i].innerHTML = '';
          cells[i].appendChild(scrollbox);
        }
        scrollbox.textContent = field.value.trim();
      } else {
        cells[i].textContent = field.value.trim();
      }
    }
  }

  // Save overrideLine status to col 15 (index 15)
  if (cells[15]) cells[15].textContent = overrideLine ? 'true' : 'false';

  // Store override status on the row itself (optional)
  if (overrideLine) {
    currentEditingRow.dataset.override = 'true';
  } else {
    delete currentEditingRow.dataset.override;
  }

  // Rerender & resort
  const tbody = document.querySelector('tbody');
  const rows = Array.from(tbody.children);

  const overrideRows = rows.filter(row => row.children[15]?.textContent.trim() === 'true');
  const autoRows = rows.filter(row => row.children[15]?.textContent.trim() !== 'true');

  autoRows.sort((a, b) => {
    const dateA = new Date(a.children[1].textContent.trim());
    const dateB = new Date(b.children[1].textContent.trim());
    if (dateA - dateB !== 0) return dateA - dateB;

    const timeA = a.children[2].textContent.trim().padStart(4, '0');
    const timeB = b.children[2].textContent.trim().padStart(4, '0');
    return parseInt(timeA, 10) - parseInt(timeB, 10);
  });

  const allRows = [...overrideRows, ...autoRows];
  allRows.forEach(row => tbody.appendChild(row));

  // Renumber: only auto-rows get new Line#
  allRows.forEach((row, index) => {
    const isOverride = row.children[15]?.textContent.trim() === 'true';
    if (!isOverride) {
      row.children[0].textContent = index + 1; // Line#
    }
    row.children[14].textContent = index + 1;    // Hidden Row#
  });
  const dropdown = document.getElementById('scriptDropdown');
  if (dropdown) dropdown.value = '';

  autoSaveTableData();
}



function saveEdits() {
      if (!currentEditingRow) {
    addNewRow(false); // treat as new entry, and close form
    return;
  }

  const cells = currentEditingRow.querySelectorAll('td');
  const newLineNum = parseInt(document.getElementById('Field1').value.trim(), 10);
  const oldLineNum = parseInt(cells[0].textContent.trim(), 10);
  const overrideLine = document.getElementById('overrideLine')?.checked;

  if (!isNaN(newLineNum) && newLineNum !== oldLineNum) {
    shiftLineNumbers(oldLineNum, newLineNum, currentEditingRow);
  }

  for (let i = 0; i < 15; i++) {
    const field = document.getElementById('Field' + (i + 1));
    if (field && cells[i]) {
      if (i === 13) {
        let scrollbox = cells[i].querySelector('.script-scrollbox');
        if (!scrollbox) {
          scrollbox = document.createElement('div');
          scrollbox.className = 'script-scrollbox';
          cells[i].innerHTML = ''; // Clear old content if needed
          cells[i].appendChild(scrollbox);
        }
        scrollbox.textContent = field.value.trim();
      } else {
        cells[i].textContent = field.value.trim();
      }

    }
  }

  if (cells[15]) {
    cells[15].textContent = overrideLine ? 'true' : 'false';
  }

  const tbody = document.querySelector('tbody');
  const rows = Array.from(tbody.children);
  const overrideRows = rows.filter(row => row.children[15]?.textContent.trim() === 'true');
  const autoRows = rows.filter(row => row.children[15]?.textContent.trim() !== 'true');

  autoRows.sort((a, b) => {
    const dateA = new Date(a.children[1].textContent.trim());
    const dateB = new Date(b.children[1].textContent.trim());

    if (dateA - dateB !== 0) return dateA - dateB;

    const timeA = a.children[2].textContent.trim().padStart(4, '0');
    const timeB = b.children[2].textContent.trim().padStart(4, '0');
    return parseInt(timeA, 10) - parseInt(timeB, 10);
  });

  const allRows = [...overrideRows, ...autoRows];
  allRows.forEach(row => tbody.appendChild(row));

  // Re-number rows
  allRows.forEach((row, index) => {
    row.children[0].textContent = index + 1;      // Line#
    row.children[14].textContent = index + 1;     // Row#
  });

  autoSaveTableData();
  cancelForm(); // ✅ Ensure this always runs
  const dropdown = document.getElementById('scriptDropdown');
  if (dropdown) dropdown.value = '';

}


function shiftLineNumbers(oldNum, newNum, excludeRow) {
  const rows = Array.from(document.querySelectorAll('tbody tr'));

  rows.forEach(row => {
    if (row === excludeRow) return;

    const cell = row.querySelector('td');
    if (!cell) return;

    const rowNum = parseInt(cell.textContent.trim(), 10);
    if (isNaN(rowNum)) return;

    if (newNum > oldNum) {
      if (rowNum > oldNum && rowNum <= newNum) {
        cell.textContent = (rowNum - 1).toString();
      }
    } else if (newNum < oldNum) {
      if (rowNum < oldNum && rowNum >= newNum) {
        cell.textContent = (rowNum + 1).toString();
      }
    }
  });
}

function addNewRow(keepOpen = true) {
  const isEditing = !!currentEditingRow;

  if (isEditing) {
    // Save existing row and reset
    if (keepOpen) {
      saveEditsOnly();
    } else {
      saveEdits();
    }
    currentEditingRow = null;
    clearForm();
    return;
  }

  const tbody = document.querySelector('tbody');
  const fields = [];
  for (let i = 1; i <= 15; i++) {
    const field = document.getElementById('Field' + i);
    fields.push(field ? field.value.trim() : '');
  }

  const time1 = fields[2].padStart(4, '0');

  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td></td>
    <td>${fields[1]}</td>
    <td>${time1}</td>
    <td>${fields[3]}</td>
    <td>${fields[4]}</td>
    <td>${fields[5]}</td>
    <td>${fields[6]}</td>
    <td>${fields[7]}</td>
    <td>${fields[8]}</td>
    <td>${fields[9]}</td>
    <td>${fields[10]}</td>
    <td>${fields[11]}</td>
    <td>${fields[12]}</td>
    <td><div class="script-scrollbox">${fields[13]}</div></td>
    <td class="hidden-col">${tbody.children.length + 1}</td>
    <td class="hidden-col">${document.getElementById('overrideLine')?.checked ? 'true' : 'false'}</td>
  `;

  tbody.appendChild(tr);
  setupRowEvents();

  sortTableByColumn(1);
  sortTableByColumn(2);

  Array.from(tbody.children).forEach((row, i) => {
    row.children[0].textContent = i + 1;
    row.children[14].textContent = i + 1;
  });

  autoSaveTableData();
  clearForm();
  currentEditingRow = null;

  if (!keepOpen) cancelForm();
}



function loadNext() {
  if (!currentEditingRow) return;
  const next = currentEditingRow.nextElementSibling;
  if (next && next.tagName === 'TR') {
    currentEditingRow = next;
    loadRowToForm(next);
  }
}

function loadPrevious() {
  if (!currentEditingRow) return;
  const prev = currentEditingRow.previousElementSibling;
  if (prev && prev.tagName === 'TR') {
    currentEditingRow = prev;
    loadRowToForm(prev);
  }
}

function loadRowToForm(row) {
  const cells = row.querySelectorAll('td');
 
  for (let i = 0; i < 15; i++) {
    const field = document.getElementById('Field' + (i + 1));
    if (field) {
      let value = cells[i].textContent.trim();

      // Auto-fix Field2 (Date)
      if (i === 1 && value) {
        const dateParts = value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
        if (dateParts) {
          const [, month, day, yearRaw] = dateParts;
          const year = yearRaw.length === 2 ? '20' + yearRaw : yearRaw;
          const monthPadded = month.padStart(2, '0');
          const dayPadded = day.padStart(2, '0');
          value = `${year}-${monthPadded}-${dayPadded}`;
        }
      }

      field.value = value;
    }
  }


  // ✅ Load override checkbox state from dataset
  document.getElementById('overrideLine').checked = cells[15]?.textContent.trim() === 'true';




  // ✅ Set override checkbox based on row dataset
  const overrideCheckbox = document.getElementById('overrideLine');
  if (overrideCheckbox) {
    overrideCheckbox.checked = row.dataset.override === 'true';
  }

  document.getElementById('injectForm').style.display = 'block';


}



function applyTemplate() { alert("Apply Template - not implemented yet."); }

function clearForm() {
  for (let i = 0; i < 15; i++) {
    const field = document.getElementById('Field' + (i + 1));
    if (field) {
      field.value = '';
    }
  }
}

function cancelForm() {
  document.getElementById('injectForm').style.display = 'none';
  document.getElementById('overlayBackground').style.display = 'none';
  clearForm();
}


// Row interaction
function setupRowEvents() {
  const rows = document.querySelectorAll('tbody tr');
  rows.forEach(row => {
    row.addEventListener('click', () => {
      document.querySelectorAll('tbody tr').forEach(r => r.classList.remove('highlighted-row'));
      row.classList.add('highlighted-row');
    });

    row.addEventListener('dblclick', () => {
      const cells = row.querySelectorAll('td');
      for (let i = 0; i < 15; i++) {  // Only load fields 1 to 15
        const field = document.getElementById('Field' + (i + 1));
        if (field) {
          field.value = cells[i].textContent.trim();
        }
      }
      currentEditingRow = row;
      document.getElementById('injectForm').style.display = 'block';
    });


  });
}

function promptTrainingUnits() {
  const count = parseInt(prompt("How many units are training?"));
  if (isNaN(count) || count <= 0) return;

  const container = document.getElementById('unitTrainingFormContainer');
  container.innerHTML = ''; // clear previous
  container.style.display = 'block';

  const form = document.createElement('form');
  form.id = 'unitTrainingForm';

  for (let i = 1; i <= count; i++) {
    const div = document.createElement('div');
    div.style.marginBottom = '0.5rem';

    div.innerHTML = `
      Unit ${i} Name: <input type="text" name="unit${i}_name" required />
      Start Date: <input type="date" name="unit${i}_start" required />
      End Date: <input type="date" name="unit${i}_end" required />
    `;

    form.appendChild(div);
  }

  const submitBtn = document.createElement('button');
  submitBtn.type = 'submit';
  submitBtn.textContent = 'Save Unit Info';
  submitBtn.style.marginTop = '1rem';
  form.appendChild(submitBtn);

  form.addEventListener('submit', handleUnitTrainingSubmit);
  container.appendChild(form);
}

function handleUnitTrainingSubmit(e) {
  e.preventDefault();

  const form = e.target;
  const display = document.getElementById('unitTrainingDisplay');
  display.innerHTML = '<h3>Unit Training Plan</h3>';

  const entries = new FormData(form);
  const grouped = {};

  for (const [key, value] of entries.entries()) {
    const match = key.match(/^unit(\d+)_(name|start|end)$/);
    if (match) {
      const [_, index, field] = match;
      if (!grouped[index]) grouped[index] = {};
      grouped[index][field] = value;
    }
  }

  const list = document.createElement('ul');
  Object.values(grouped).forEach(unit => {
    const li = document.createElement('li');
    li.textContent = `${unit.name} — Training from ${unit.start} to ${unit.end}`;
    list.appendChild(li);
  });

  display.appendChild(list);
  localStorage.setItem('unitTrainingData', JSON.stringify(grouped));
  document.getElementById('unitTrainingFormContainer').style.display = 'none';
}

function loadSavedUnitTraining() {
  const saved = localStorage.getItem('unitTrainingData');
  if (!saved) return;

  const grouped = JSON.parse(saved);
  const display = document.getElementById('unitTrainingDisplay');
  display.innerHTML = '<h3>Unit Training Plan</h3>';
  const list = document.createElement('ul');

  Object.values(grouped).forEach(unit => {
    const li = document.createElement('li');
    li.textContent = `${unit.name} — Training from ${unit.start} to ${unit.end}`;
    list.appendChild(li);
  });

  display.appendChild(list);
}


function autoSaveTableData() {
  const rows = document.querySelectorAll('tbody tr');
  const data = [];

  rows.forEach(row => {
    const cells = row.querySelectorAll('td');
    const rowData = [];
    cells.forEach(cell => rowData.push(cell.textContent.trim()));
    data.push(rowData);
  });

  localStorage.setItem('mselInjectData', JSON.stringify(data));
  const unitTraining = localStorage.getItem('unitTrainingData');
  const scripts = localStorage.getItem('scriptLibrary') || '{}';

}

async function manualBackupDownload() {
  const rows = document.querySelectorAll('tbody tr');
  const data = [];

  rows.forEach(row => {
    const cells = row.querySelectorAll('td');
    const rowData = [];
    cells.forEach(cell => rowData.push(cell.textContent.trim()));
    data.push(rowData);
  });

  const unitTraining = localStorage.getItem('unitTrainingData');
  const scripts = localStorage.getItem('scriptLibrary') || '{}';

  const backup = {
    injects: data,
    unitTraining: unitTraining ? JSON.parse(unitTraining) : {},
    scriptLibrary: JSON.parse(scripts)
  };

  // Ask user where to save
  try {
    const opts = {
      suggestedName: "MSEL_Backup.json",
      types: [{
        description: "JSON Files",
        accept: { "application/json": [".json"] }
      }]
    };

    const handle = await window.showSaveFilePicker(opts);
    const writable = await handle.createWritable();
    await writable.write(JSON.stringify(backup, null, 2));
    await writable.close();
    alert("Backup saved successfully!");
  } catch (err) {
    console.warn("Save canceled or failed:", err);
  }
}



// Save every 60 seconds
setInterval(autoSaveTableData, 60000);

function clearSavedData() {
  if (confirm("Are you sure you want to clear saved MSEL data?")) {
    localStorage.removeItem('mselInjectData');
    localStorage.removeItem('unitTrainingData');
    location.reload();
  }
}

let currentScriptKey = null;

function showScriptEditor() {
  document.getElementById('scriptEditorContainer').style.display = 'block';
  loadScriptList();
  clearScriptForm();
}

function loadScriptList() {
  const list = document.getElementById('scriptList');
  list.innerHTML = '';

  const scripts = JSON.parse(localStorage.getItem('scriptsData') || localStorage.getItem('scriptLibrary') || '{}');
  Object.keys(scripts).forEach(title => {
    const li = document.createElement('li');
    li.textContent = title;
    li.style.cursor = 'pointer';
    li.style.padding = '5px';
    li.style.borderBottom = '1px solid #eee';
    li.onclick = () => {
      document.getElementById('scriptTitle').value = title;
      document.getElementById('scriptBody').value = scripts[title];
      currentScriptKey = title;
    };
    list.appendChild(li);
  });
}

function saveScript() {
  const title = document.getElementById('scriptTitle').value.trim();
  const body = document.getElementById('scriptBody').value;

  if (!title) return alert("Title is required.");

  const scripts = JSON.parse(localStorage.getItem('scriptsData') || localStorage.getItem('scriptLibrary') || '{}');

  // Rename logic
  if (currentScriptKey && currentScriptKey !== title) {
    // Delete the old script name
    delete scripts[currentScriptKey];
  }

  // Save new or updated script
  scripts[title] = body;

  // Save to both keys
  localStorage.setItem('scriptsData', JSON.stringify(scripts));
  localStorage.setItem('scriptLibrary', JSON.stringify(scripts)); // keep backup consistent

  currentScriptKey = title; // update tracking key
  loadScriptList();
  populateScriptDropdown();
  alert("Script saved.");
}


function deleteScript() {
  if (!currentScriptKey) return;
  const scripts = JSON.parse(localStorage.getItem('scriptsData') || localStorage.getItem('scriptLibrary') || '{}');
  if (scripts[currentScriptKey]) {
    delete scripts[currentScriptKey];
    localStorage.setItem('scriptsData', JSON.stringify(scripts));
    loadScriptList();
    clearScriptForm();
    alert("Script deleted.");
  }
}

function cancelScriptEdit() {
  clearScriptForm();
  document.getElementById('scriptEditorContainer').style.display = 'none';
}

function newScript() {
  clearScriptForm();
  currentScriptKey = null;
}

function clearScriptForm() {
  document.getElementById('scriptTitle').value = '';
  document.getElementById('scriptBody').value = '';
}


function populateScriptDropdown() {
  const dropdown = document.getElementById('scriptDropdown');
  if (!dropdown) return;

  dropdown.innerHTML = '';

  const defaultOption = document.createElement('option');
  defaultOption.value = '';
  defaultOption.textContent = '-- Select a Script --';
  dropdown.appendChild(defaultOption);

  const scripts = JSON.parse(localStorage.getItem('scriptsData') || localStorage.getItem('scriptLibrary') || '{}');
  Object.keys(scripts).forEach(title => {
    const option = document.createElement('option');
    option.value = title;
    option.textContent = title;
    dropdown.appendChild(option);
  });
}


function loadFromBackup() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';

  input.addEventListener('change', event => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const backup = JSON.parse(e.target.result);
        if (backup.injects) {
          renderTable(backup.injects);
          localStorage.setItem('mselInjectData', JSON.stringify(backup.injects));
        }
        if (backup.unitTraining) {
          localStorage.setItem('unitTrainingData', JSON.stringify(backup.unitTraining));
          loadSavedUnitTraining();
        }
        if (backup.scriptLibrary) {
          localStorage.setItem('scriptLibrary', JSON.stringify(backup.scriptLibrary));
          populateScriptDropdown(); 
        }
        alert('Backup loaded successfully!');
      } catch (err) {
        alert('Error loading backup file.');
        console.error(err);
      }
    };

    reader.readAsText(file);
  });
  resizeProxyWidth(); // ← ensures scroll proxy matches actual table width
  manageScrollVisibility();

  input.click();
}

async function saveBackupWithDialog(backupData) {
  try {
    const opts = {
      suggestedName: "MSEL_Backup.json",
      types: [
        {
          description: "JSON Files",
          accept: { "application/json": [".json"] }
        }
      ]
    };

    const handle = await window.showSaveFilePicker(opts);
    const writable = await handle.createWritable();
    await writable.write(JSON.stringify(backupData, null, 2));
    await writable.close();
    alert("Backup saved successfully!");
  } catch (err) {
    console.error("Save canceled or failed", err);
  }
}

  /*const unitData = JSON.parse(localStorage.getItem('unitTrainingData') || '{}');
  const injectRows = document.querySelectorAll('tbody tr');
  const sheetData = [];

  // Title row
  sheetData.push(["MSEL Inject Sheet"]);
  sheetData.push([]);

  // Unit Training Info
  sheetData.push(["Unit Training Plan"]);
  for (const key in unitData) {
    const unit = unitData[key];
    sheetData.push([`${unit.name} — Training from ${unit.start} to ${unit.end}`]);
  }
  sheetData.push([]);

  // Table headers
  const headers = [
    'Line #', 'Date', 'Time', 'Storyline', 'Event', 'Inject', 'Description',
    'Control Zone', 'Inject Platform', 'Method',
    'Target Audience', 'Narrative for storyline / Notes', 'T&R Event'
  ];
  sheetData.push(headers);

  // Inject rows
  injectRows.forEach(row => {
    const cells = row.querySelectorAll('td');
    const rowData = [];
    for (let i = 0; i < 13; i++) {
      rowData.push(cells[i]?.textContent.trim() || '');
    }
    sheetData.push(rowData);
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "MSEL");

  // Merge A1:N1 for the title row
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];

  // Autosize columns
  ws['!cols'] = headers.map((_, i) => {
    let maxLen = headers[i].length;
    injectRows.forEach(row => {
      const txt = row.querySelectorAll('td')[i]?.textContent || '';
      maxLen = Math.max(maxLen, txt.length);
    });
    return { wch: maxLen + 2 };
  });

  // Style the cells
  const range = XLSX.utils.decode_range(ws['!ref']);
  const headerRowIndex = sheetData.findIndex(r => r[0] === 'Line #');
  for (let R = range.s.r; R <= range.e.r; ++R) {
    const isHeader = R === headerRowIndex;
    const isAltRow = (R - headerRowIndex) % 2 === 1;

    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellRef]) continue;

      const cell = ws[cellRef];
      cell.s = {
        font: {},
        alignment: { vertical: "center", wrapText: true },
        border: {
          top:    { style: isHeader ? "medium" : "thin" },
          bottom: { style: isHeader ? "medium" : "thin" },
          left:   { style: isHeader ? "medium" : "thin" },
          right:  { style: isHeader ? "medium" : "thin" }
        },
        fill: {}
      };

      if (R === 0) {
        cell.s.font = { bold: true, sz: 16 };
        cell.s.alignment.horizontal = "center";
      }

      if (isHeader) {
        cell.s.font = { bold: true };
        cell.s.alignment.horizontal = "center";
        cell.s.fill = { fgColor: { rgb: "D9E1F2" } }; // Light blue
      } else if (R > headerRowIndex && isAltRow) {
        cell.s.fill = { fgColor: { rgb: "F2F2F2" } }; // Light gray for alternating rows
      }
    }
  }

  const firstUnitName = Object.values(unitData)[0]?.name?.replace(/\s+/g, '_') || 'Unit';
  const today = new Date().toISOString().split('T')[0];
  const filename = `${firstUnitName}_${today}_Formatted.xlsx`;

  XLSX.writeFile(wb, filename, { cellStyles: true });*/
 async function exportToExcel() {
  const unitData = JSON.parse(localStorage.getItem('unitTrainingData') || '{}');
  const injectRows = document.querySelectorAll('tbody tr');

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("MSEL");

  // Title Row
  worksheet.mergeCells('A1:M1');
  const titleCell = worksheet.getCell('A1');
  titleCell.value = 'MSEL Inject Sheet';
  titleCell.font = { bold: true, size: 16 };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  let rowIdx = 3;

  // Unit Training Info (Merged across A to D)
  if (Object.keys(unitData).length > 0) {
    worksheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
    const title = worksheet.getCell(`A${rowIdx}`);
    title.value = 'Unit Training Plan';
    title.font = { bold: true };
    rowIdx++;

    for (const key in unitData) {
      const unit = unitData[key];
      worksheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      worksheet.getCell(`A${rowIdx}`).value = `${unit.name} — Training from ${unit.start} to ${unit.end}`;
      rowIdx++;
    }
    rowIdx++; // spacer row
  }

  // Headers
  const headers = [
    'Line #', 'Date', 'Time', 'Storyline', 'Event', 'Inject', 'Description',
    'Control Zone', 'Inject Platform', 'Method',
    'Target Audience', 'Narrative for storyline / Notes', 'T&R Event'
  ];

  const headerRow = worksheet.getRow(rowIdx);
  headerRow.values = headers;

  const centerCols = [1, 2, 3, 4, 5, 6, 8, 9, 10, 11]; // 1-based indices to center-align

  headerRow.eachCell((cell, colNumber) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'EEEEEE' } // neutral light gray
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: centerCols.includes(colNumber) ? 'center' : 'left',
      wrapText: true
    };
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    };
  });

  rowIdx++;

  // Inject Data Rows
  injectRows.forEach((row, i) => {
    const cells = row.querySelectorAll('td');
    const rowData = [];
    for (let j = 0; j < 13; j++) {
      rowData.push(cells[j]?.textContent.trim() || '');
    }
    const newRow = worksheet.insertRow(rowIdx++, rowData);
    const isAlt = i % 2 === 1;

    newRow.eachCell(cell => {
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      if (isAlt) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'F9F9F9' }
        };
      }
    });
  });

  // Autosize Columns
  worksheet.columns.forEach((col, i) => {
    let maxLength = headers[i].length;
    col.eachCell({ includeEmpty: true }, cell => {
      const val = cell.value ? String(cell.value) : '';
      maxLength = Math.max(maxLength, val.length);
    });
    col.width = maxLength + 2;
  });

  // Save to file
  const firstUnitName = Object.values(unitData)[0]?.name?.replace(/\s+/g, '_') || 'Unit';
  const today = new Date().toISOString().split('T')[0];
  const filename = `${firstUnitName}_${today}_Formatted.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}




async function exportToPDF() {
  const { jsPDF } = window.jspdf;

  // Prompt for paper size selection BEFORE doing anything else
  const selectedSize = prompt("Select paper size: Type 'letter' or 'legal'", "letter");
  if (!selectedSize || !['letter', 'legal'].includes(selectedSize.toLowerCase())) {
    alert("Export canceled or invalid input.");
    return;
  }
  const format = selectedSize.toLowerCase();

  const doc = new jsPDF({
    orientation: "landscape",
    unit: "pt",
    format: format
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const startY = 50;

  // Title
  doc.setFont("helvetica", "bold");
  doc.setFontSize(20);
  doc.text("MSEL SCRIPT", pageWidth / 2, startY, { align: "center" });

  let yOffset = startY + 30;

  // Add unit info
  const unitData = JSON.parse(localStorage.getItem('unitTrainingData') || '{}');
  const unitInfoLines = Object.values(unitData).map(unit =>
    `${unit.name} — Training from ${unit.start} to ${unit.end}`
  );

  doc.setFont("helvetica", "normal");
  doc.setFontSize(12);
  unitInfoLines.forEach((line, index) => {
    doc.text(line, 40, yOffset + (index * 18));
  });

  yOffset += (unitInfoLines.length + 1) * 18;

  // Table headers and data
  const headers = [[
    'Line #', 'Date', 'Time', 'Storyline', 'Event', 'Inject', 'Description',
    'Control Zone', 'Inject Platform', 'Method',
    'Target Audience', 'Narrative / Notes', 'T&R Event'
  ]];

  const rows = Array.from(document.querySelectorAll("tbody tr")).map(tr => {
    const cells = tr.querySelectorAll("td");
    const rowData = [];
    for (let i = 0; i < 13; i++) {
      rowData.push(cells[i]?.textContent.trim() || "");
    }
    return rowData;
  });

doc.autoTable({
  startY: yOffset,
  head: headers,
  body: rows,
  styles: {
    fontSize: 9,            // Slightly larger text
    cellPadding: 2,         // Tighter padding
  },
  headStyles: {
    fillColor: [30, 30, 30],
    textColor: [255, 255, 255],
    halign: "center",
    fontStyle: "bold"
  },
  alternateRowStyles: {
    fillColor: [245, 245, 245]
  },
  margin: { left: 20, right: 20 },  // Reduced margins
  tableWidth: 'auto'
});

  const firstUnitName = Object.values(unitData)[0]?.name?.replace(/\s+/g, "_") || "Unit";
  const today = new Date().toISOString().split("T")[0];
  const filename = `${firstUnitName}_${today}_MSEL.pdf`;

  // Ask user where to save using File System Access API (if supported)
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [
          {
            description: 'PDF Document',
            accept: { 'application/pdf': ['.pdf'] }
          }
        ]
      });

      const writable = await handle.createWritable();
      const pdfBlob = doc.output('blob');
      await writable.write(pdfBlob);
      await writable.close();
      alert("PDF exported successfully.");
    } catch (err) {
      console.warn("Save canceled or failed:", err);
    }
  } else {
    // Fallback for browsers that don't support showSaveFilePicker
    doc.save(filename);
  }
}

async function exportInjectsToPDF() {
  const { jsPDF } = window.jspdf;

  const unitData = JSON.parse(localStorage.getItem('unitTrainingData') || '{}');
  const injectRows = document.querySelectorAll('tbody tr');

  const unitLines = Object.values(unitData).map(
    unit => `${unit.name} — Training from ${unit.start} to ${unit.end}`
  );

  const doc = new jsPDF({
    orientation: "portrait",
    unit: "pt",
    format: "letter"
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 40;
  const lineHeight = 16;

  injectRows.forEach((row, index) => {
    if (index > 0) doc.addPage();

    const cells = row.querySelectorAll('td');
    let y = margin;

    // === HEADER ===
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("UNCLASSIFIED", pageWidth / 2, y, { align: "center" });
    y += lineHeight + 4;

    doc.setFontSize(18);
    doc.text("MSEL Scripts", pageWidth / 2, y, { align: "center" });
    y += lineHeight + 2;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    // unitLines.forEach((line, i) => {
    //   doc.text(line, pageWidth / 2, y + (i * lineHeight), { align: "center" });
    // });

    // y += unitLines.length * lineHeight + 10;

    // === FIELDS ===
    const fields = [
      ['Line #', cells[0]?.textContent],
      ['Date', cells[1]?.textContent],
      ['Time', cells[2]?.textContent],
      ['Storyline', cells[3]?.textContent],
      ['Event', cells[4]?.textContent],
      ['Description', cells[6]?.textContent],
      ['Narrative / Notes', cells[11]?.textContent],
      ['T&R Event', cells[12]?.textContent],
    ];

    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);

    for (const [label, value] of fields) {
      if (value && value.trim()) {
        const wrapped = doc.splitTextToSize(`${label}: ${value.trim()}`, pageWidth - 2 * margin);
        if (y + wrapped.length * lineHeight > pageHeight - margin) {
          doc.addPage();
          y = margin;
        }
        doc.text(wrapped, margin, y);
        y += wrapped.length * lineHeight + 4;
      }
    }

    // === SCRIPT BLOCK ===
    const scriptText = cells[13]?.textContent?.trim();
    if (scriptText) {
      const scriptLabel = "Script:";
      const wrappedScript = doc.splitTextToSize(scriptText, pageWidth - 2 * margin);

      if (y + lineHeight + wrappedScript.length * lineHeight > pageHeight - margin) {
        doc.addPage();
        y = margin;
      }

      doc.setFont("helvetica", "bold");
      doc.text(scriptLabel, margin, y);
      y += lineHeight;

      doc.setFont("helvetica", "normal");
      while (wrappedScript.length > 0) {
        if (y + lineHeight > pageHeight - margin) {
          doc.addPage();
          y = margin;
        }
        const chunk = wrappedScript.splice(0, Math.floor((pageHeight - y - margin) / lineHeight));
        doc.text(chunk, margin, y);
        y += chunk.length * lineHeight;
      }
    }
  });

  const unitName = Object.values(unitData)[0]?.name?.replace(/\s+/g, '_') || 'Unit';
  const today = new Date().toISOString().split('T')[0];
  const filename = `${unitName}_${today}_Scripts_ByPage.pdf`;

  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: filename,
      types: [{
        description: 'PDF Document',
        accept: { 'application/pdf': ['.pdf'] }
      }]
    });
    const writable = await handle.createWritable();
    const pdfBlob = doc.output('blob');
    await writable.write(pdfBlob);
    await writable.close();
    alert("PDF saved successfully.");
  } catch (err) {
    console.warn("Save canceled or failed:", err);
  }
}

function openInjectForm() {
  document.getElementById('injectForm').style.display = 'block';
  document.getElementById('overlayBackground').style.display = 'block';
  clearForm();
  currentEditingRow = null;
  // Ensure this runs after all Field1–Field14 are generated and added to .form-fields
const formFieldsContainer = document.querySelector('.form-fields');

// Avoid duplicate checkbox on repeated opens
if (!document.getElementById('overrideLine')) {
  const overrideDiv = document.createElement('div');
  overrideDiv.style.width = '100%';
  overrideDiv.style.marginTop = '1rem';
  overrideDiv.innerHTML = `
    <label style="font-weight: bold;">
      <input type="checkbox" id="overrideLine" />
      Override Auto Line Number
    </label>
  `;
  formFieldsContainer.appendChild(overrideDiv);
}

}

function exportInjectsToDOCX() {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak } = window.docx;

  const unitData = JSON.parse(localStorage.getItem('unitTrainingData') || '{}');
  const injectRows = document.querySelectorAll('tbody tr');

  const unitInfoLines = Object.values(unitData).map(
    unit => `${unit.name} — Training from ${unit.start} to ${unit.end}`
  );

  const docSections = [];

  injectRows.forEach((row, index) => {
    const cells = row.querySelectorAll('td');

    const injectFields = [
      ['Line #', cells[0]?.textContent || ''],
      ['Date', cells[1]?.textContent || ''],
      ['Time', cells[2]?.textContent || ''],
      ['Storyline', cells[3]?.textContent || ''],
      ['Event', cells[4]?.textContent || ''],
      ['Description', cells[6]?.textContent || ''],
      ['Narrative / Notes', cells[11]?.textContent || ''],
      ['T&R Event', cells[12]?.textContent || '']
    ];

    const scriptText = cells[13]?.textContent?.trim() || '';

    const headerParagraphs = [
    new Paragraph({
      children: [
        new TextRun({
          text: 'UNCLASSIFIED',
          bold: true,
          color: '00AA00' // green hex code
        })
      ],
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER
    }),

      new Paragraph({
        text: 'MSEL Scripts',
        heading: HeadingLevel.HEADING_2,
        alignment: AlignmentType.CENTER
      }),
      //...unitInfoLines.map(line =>
      //  new Paragraph({
      //    text: line,
      //    alignment: AlignmentType.CENTER
      //  })
      //),
      new Paragraph({ text: '' }) // spacer
    ];

    const fieldParagraphs = injectFields.flatMap(([label, value]) => {
      if (!value.trim()) return [];
      return [
        new Paragraph({
          children: [
            new TextRun({ text: `${label}: `, bold: true }),
            new TextRun({ text: value })
          ]
        })
      ];
    });

    const scriptParagraphs = [
      new Paragraph({ text: '' }),
      new Paragraph({ text: 'Script:', bold: true }),
      ...scriptText.split('\n').map(line => new Paragraph({ text: line }))
    ];

    const pageParagraphs = [
      ...headerParagraphs,
      ...fieldParagraphs,
      ...scriptParagraphs,
      new Paragraph({ children: [new PageBreak()] })
    ];

    docSections.push(...pageParagraphs);
  });

  const doc = new Document({
    sections: [{ properties: {}, children: docSections }]
  });

  Packer.toBlob(doc).then(blob => {
    const unitName = Object.values(unitData)[0]?.name?.replace(/\s+/g, '_') || 'Unit';
    const today = new Date().toISOString().split('T')[0];
    const filename = `${unitName}_${today}_Scripts_ByPage.docx`;

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
  });
}

const tableScroll = document.getElementById('tableScrollWrapper');
const proxyScroll = document.getElementById('horizontalScrollProxy');

function syncScrolls() {
  proxyScroll.scrollLeft = tableScroll.scrollLeft;
}
function syncProxyScroll() {
  tableScroll.scrollLeft = proxyScroll.scrollLeft;
}

proxyScroll.addEventListener('scroll', syncProxyScroll);
tableScroll.addEventListener('scroll', syncScrolls);

// Show/hide proxy depending on scroll position
function manageScrollVisibility() {
  const tableBottom = tableScroll.getBoundingClientRect().bottom;
  const windowBottom = window.innerHeight;

  if (tableBottom < windowBottom - 30) {
    proxyScroll.style.display = 'block';
  } else {
    proxyScroll.style.display = 'none';
  }
}

function resizeProxyWidth() {
  const table = document.querySelector('table');
  const proxy = document.getElementById('scrollProxyInner');
  if (table && proxy) {
    proxy.style.width = table.scrollWidth + 'px';
  }
}


window.addEventListener('scroll', manageScrollVisibility);
window.addEventListener('resize', () => {
  resizeProxyWidth();
  manageScrollVisibility();
});


function syncProxyWidthToTable() {
  const tableWrapper = document.getElementById('tableScrollWrapper');
  const proxyInner = document.getElementById('scrollProxyInner');
  if (!tableWrapper || !proxyInner) return;

  // Sync proxy width to the scroll width of the real wrapper
  proxyInner.style.width = tableWrapper.scrollWidth + 'px';
}

function initScriptCellBehavior() {
  const cells = document.querySelectorAll('textarea.script-cell');

  cells.forEach(cell => {
    cell.style.overflowY = 'hidden';
    cell.style.resize = 'vertical';
    cell.style.minHeight = '1.6em';
    cell.style.height = 'auto';
    cell.style.maxHeight = 'none';
    cell.style.boxSizing = 'border-box';
    cell.style.fontFamily = 'inherit';
    cell.style.fontSize = 'inherit';

    // Shrink to content height, but not less than 1.6em
    requestAnimationFrame(() => {
      cell.style.height = Math.max(cell.scrollHeight, cell.offsetHeight, 24) + 'px';
    });

    // Optional: auto-expand on input
    cell.addEventListener('input', () => {
      cell.style.height = 'auto';
      cell.style.height = Math.max(cell.scrollHeight, 24) + 'px';
    });
  });
}



  // Make available to button onclick
  window.exportToExcel = exportToExcel;

  function exportScripts() {
  const scripts = JSON.parse(localStorage.getItem('scriptsData') || '{}');
  const blob = new Blob([JSON.stringify(scripts, null, 2)], { type: 'application/json' });

  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'scripts_backup.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function importScripts() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';

  input.addEventListener('change', event => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const newScripts = JSON.parse(e.target.result);
        if (typeof newScripts !== 'object' || Array.isArray(newScripts)) {
          alert('Invalid script format.');
          return;
        }

        const current = JSON.parse(localStorage.getItem('scriptsData') || '{}');
        const merged = { ...current, ...newScripts };
        localStorage.setItem('scriptsData', JSON.stringify(merged));
        localStorage.setItem('scriptLibrary', JSON.stringify(merged));
        populateScriptDropdown();
        loadScriptList();
        alert('Scripts imported successfully!');
      } catch (err) {
        alert('Failed to import scripts.');
        console.error(err);
      }
    };
    reader.readAsText(file);
  });

  input.click();
}

function deleteCurrentRow() {
  if (!currentEditingRow) return alert("No record selected.");

  const confirmed = confirm("Are you sure you want to delete this record?");
  if (!confirmed) return;

  currentEditingRow.remove();
  currentEditingRow = null;

  // Re-index table rows
  const tbody = document.querySelector('tbody');
  Array.from(tbody.children).forEach((row, index) => {
    row.children[0].textContent = index + 1;  // Line #
    row.children[14].textContent = index + 1; // Row #
  });

  autoSaveTableData();
  cancelForm();
}
