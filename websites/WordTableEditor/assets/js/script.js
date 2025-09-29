
const input = document.getElementById('fileInput');
const tablesContainer = document.getElementById('tables');
const downloadBtn = document.getElementById('downloadJson');
const downloadExcelBtn = document.getElementById('downloadExcel');
const uploadBtn = document.getElementById('uploadServer');
const textToolbar = document.getElementById('textToolbar');
const fontNameSelect = document.getElementById('fontName');
const fontSizeSelect = document.getElementById('fontSize');
const closeToolbarBtn = document.querySelector('.close-toolbar');
const colorPickerBtn = document.getElementById('colorPickerBtn');
const colorPalette = document.getElementById('colorPalette');
const exportModal = document.getElementById('exportModal');
const tableList = document.getElementById('tableList');
const selectAllCheckbox = document.getElementById('selectAllCheckbox');
const cancelExportBtn = document.getElementById('cancelExport');
const confirmExportBtn = document.getElementById('confirmExport');
const languageBtn = document.getElementById('languageBtn');
const currentLanguageElem = document.getElementById('currentLanguage');
const instructionModal = document.getElementById('instructionModal');
const openInstructionBtn = document.getElementById('openInstruction');
const closeInstructionBtn = document.getElementById('closeInstruction');


let currentLanguage = 'ru';
let extractedData = [];
let activeCell = null;
let isSelectingColor = false;
let selectedTablesForExport = [];
let activeTableIndex = null;
let activeColumnIndex = null;
let activeRowIndex = null;

// Функция для установки языка
function setLanguage(lang) {
  currentLanguage = lang;
  currentLanguageElem.textContent = lang.toUpperCase();

  // Обновляем все элементы с атрибутом data-translate
  document.querySelectorAll('[data-translate]').forEach(element => {
    const key = element.getAttribute('data-translate');
    if (translations[lang][key]) {
      element.textContent = translations[lang][key];
    }
  });

  // Обновляем title атрибуты
  document.querySelectorAll('[data-translate-title]').forEach(element => {
    const key = element.getAttribute('data-translate-title');
    if (translations[lang][key]) {
      element.setAttribute('title', translations[lang][key]);
    }
  });

  // Перерисовываем таблицы для обновления текста кнопок
  if (extractedData.length > 0) {
    renderTables();
  }
}

// Переключение языка
languageBtn.addEventListener('click', () => {
  const newLang = currentLanguage === 'ru' ? 'az' : 'ru';
  setLanguage(newLang);
});

// Открытие инструкции
openInstructionBtn.addEventListener('click', () => {
  instructionModal.classList.add('visible');
});

// Закрытие инструкции
closeInstructionBtn.addEventListener('click', () => {
  instructionModal.classList.remove('visible');
});

// Функция для экспорта в Excel с выбранными таблицами
function exportToExcel(tablesToExport = null) {
  if (extractedData.length === 0) {
    alert(currentLanguage === 'ru' ? 'Нет данных для экспорта!' : 'İxrac üçün məlumat yoxdur!');
    return;
  }

  const tables = tablesToExport || extractedData;

  if (tables.length === 0) {
    alert(currentLanguage === 'ru' ? 'Не выбрано ни одной таблицы для экспорта!' : 'İxrac üçün heç bir cədvəl seçilməyib!');
    return;
  }

  try {
    const wb = XLSX.utils.book_new();

    tables.forEach((tableData, index) => {
      const excelData = [];

      excelData.push([tableData.title]);
      excelData.push([]);

      if (tableData.table && tableData.table.length > 0) {
        tableData.table.forEach(row => {
          const cleanRow = row.map(cell => {
            const div = document.createElement('div');
            div.innerHTML = cell;
            return div.textContent || div.innerText || '';
          });
          excelData.push(cleanRow);
        });
      }

      const ws = XLSX.utils.aoa_to_sheet(excelData);

      if (!ws['A1']) ws['A1'] = { t: 's' };
      ws['A1'].s = {
        font: { bold: true, sz: 14 },
        alignment: { horizontal: 'center' }
      };

      const colWidths = [];
      if (tableData.table && tableData.table.length > 0) {
        const firstDataRow = tableData.table[0];
        for (let i = 0; i < firstDataRow.length; i++) {
          colWidths.push({ wch: 20 });
        }
        ws['!cols'] = colWidths;
      }

      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell_address = { c: C, r: R };
          const cell_ref = XLSX.utils.encode_cell(cell_address);

          if (!ws[cell_ref]) continue;

          if (!ws[cell_ref].s) ws[cell_ref].s = {};
          ws[cell_ref].s.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          };

          if (R === 2 && tableData.table && tableData.table[0]) {
            if (!ws[cell_ref].s.font) ws[cell_ref].s.font = {};
            ws[cell_ref].s.font.bold = true;
            ws[cell_ref].s.fill = { fgColor: { rgb: "E6E6FA" } };
          }
        }
      }

      let sheetName = tableData.title.replace(/[\\/*?:[\]]/g, '').substring(0, 31);
      if (sheetName === '') sheetName = `${currentLanguage === 'ru' ? 'Таблица' : 'Cədvəl'} ${index + 1}`;

      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    const fileName = `lesson_plan_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);

  } catch (error) {
    console.error('Ошибка при экспорте в Excel:', error);
    alert((currentLanguage === 'ru' ? 'Произошла ошибка при создании Excel файла: ' : 'Excel faylı yaradılarkən xəta baş verdi: ') + error.message);
  }
}

// Показать модальное окно выбора таблиц
function showExportModal() {
  tableList.innerHTML = '';

  extractedData.forEach((table, index) => {
    const listItem = document.createElement('div');
    listItem.className = 'table-list-item';
    listItem.innerHTML = `
      <input type="checkbox" class="table-checkbox" id="modal-table-${index}" data-index="${index}" checked>
      <label for="modal-table-${index}" class="flex-1">${table.title} (ID: ${index})</label>
    `;
    tableList.appendChild(listItem);
  });

  selectAllCheckbox.checked = true;
  exportModal.classList.add('visible');
}

// Обновление заголовка таблицы
function updateTableTitle(index, newTitle) {
  if (extractedData[index]) {
    extractedData[index].title = newTitle;
    saveTableContent();
  }
}

// Удаление таблицы
function deleteTable(index) {
  const confirmMessage = currentLanguage === 'ru'
    ? 'Вы уверены, что хотите удалить эту таблицу?'
    : 'Bu cədvəli silmək istədiyinizə əminsiniz?';

  if (confirm(confirmMessage)) {
    extractedData.splice(index, 1);
    renderTables();
    saveTableContent();

    if (extractedData.length === 0) {
      downloadBtn.disabled = true;
      downloadExcelBtn.disabled = true;
      uploadBtn.disabled = true;
    }
  }
}

// Добавление пустого столбца
function addColumn(tableIndex, position = 'end') {
  if (extractedData[tableIndex]) {
    // Сохраняем текущие данные перед изменением структуры
    saveTableContent();

    const table = extractedData[tableIndex];

    if (table.table && table.table.length > 0) {
      table.table.forEach(row => {
        if (position === 'start') {
          row.unshift('');
        } else {
          row.push('');
        }
      });
    }

    renderTables();
  }
}

// Удаление пустых столбцов
function removeEmptyColumns(tableIndex) {
  if (extractedData[tableIndex]) {
    // Сохраняем текущие данные перед изменением структуры
    saveTableContent();

    const table = extractedData[tableIndex];

    if (table.table && table.table.length > 0) {
      const emptyColumns = [];
      const colCount = table.table[0].length;

      for (let col = 0; col < colCount; col++) {
        const isEmpty = table.table.every(row => {
          const cellContent = row[col] || '';
          return cellContent.trim() === '';
        });

        if (isEmpty) {
          emptyColumns.push(col);
        }
      }

      emptyColumns.reverse().forEach(colIndex => {
        table.table.forEach(row => {
          row.splice(colIndex, 1);
        });
      });

      if (emptyColumns.length > 0) {
        alert(currentLanguage === 'ru'
          ? `Удалено ${emptyColumns.length} пустых столбцов.`
          : `${emptyColumns.length} boş sütun silindi.`);
      } else {
        alert(currentLanguage === 'ru'
          ? 'Пустых столбцов не найдено.'
          : 'Boş sütun tapılmadı.');
      }
    }

    renderTables();
  }
}

// Удаление текущего столбца
function removeCurrentColumn(tableIndex, columnIndex) {
  if (extractedData[tableIndex] && extractedData[tableIndex].table[0] && columnIndex >= 0) {
    const confirmMessage = currentLanguage === 'ru'
      ? 'Вы уверены, что хотите удалить этот столбец?'
      : 'Bu sütunu silmək istədiyinizə əminsiniz?';

    if (confirm(confirmMessage)) {
      // Сохраняем текущие данные перед изменением структуры
      saveTableContent();

      extractedData[tableIndex].table.forEach(row => {
        row.splice(columnIndex, 1);
      });

      renderTables();
    }
  }
}

// Добавление пустой строки
function addRow(tableIndex, position = 'end') {
  if (extractedData[tableIndex]) {
    // Сохраняем текущие данные перед изменением структуры
    saveTableContent();

    const table = extractedData[tableIndex];

    if (table.table && table.table.length > 0) {
      const colCount = table.table[0].length;
      const newRow = Array(colCount).fill('');

      if (position === 'start') {
        table.table.unshift(newRow);
      } else {
        table.table.push(newRow);
      }
    } else {
      table.table = [Array(1).fill('')];
    }

    renderTables();
  }
}

// Удаление пустых строк
function removeEmptyRows(tableIndex) {
  if (extractedData[tableIndex]) {
    // Сохраняем текущие данные перед изменением структуры
    saveTableContent();

    const table = extractedData[tableIndex];

    if (table.table && table.table.length > 0) {
      const emptyRows = [];

      table.table.forEach((row, rowIndex) => {
        const isEmpty = row.every(cell => {
          const cellContent = cell || '';
          return cellContent.trim() === '';
        });

        if (isEmpty) {
          emptyRows.push(rowIndex);
        }
      });

      emptyRows.reverse().forEach(rowIndex => {
        table.table.splice(rowIndex, 1);
      });

      if (emptyRows.length > 0) {
        alert(currentLanguage === 'ru'
          ? `Удалено ${emptyRows.length} пустых строк.`
          : `${emptyRows.length} boş sətir silindi.`);
      } else {
        alert(currentLanguage === 'ru'
          ? 'Пустых строк не найдено.'
          : 'Boş sətir tapılmadı.');
      }
    }

    renderTables();
  }
}

// Удаление текущей строки
function removeCurrentRow(tableIndex, rowIndex) {
  if (extractedData[tableIndex] && extractedData[tableIndex].table[rowIndex]) {
    const confirmMessage = currentLanguage === 'ru'
      ? 'Вы уверены, что хотите удалить эту строку?'
      : 'Bu sətri silmək istədiyinizə əminsiniz?';

    if (confirm(confirmMessage)) {
      // Сохраняем текущие данные перед изменением структуры
      saveTableContent();

      extractedData[tableIndex].table.splice(rowIndex, 1);
      renderTables();
    }
  }
}

// Перерисовка таблиц в интерфейсе
function renderTables() {
  tablesContainer.innerHTML = '';

  extractedData.forEach((tableData, idx) => {
    const editableTable = document.createElement('table');
    if (tableData.table && tableData.table.length > 0) {
      tableData.table.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');

        // Добавляем номер строки
        const rowNumberCell = document.createElement('td');
        rowNumberCell.innerHTML = `<span class="row-number">${rowIndex + 1}</span>`;
        rowNumberCell.style.width = '30px';
        rowNumberCell.style.background = '#f3f4f6';
        rowNumberCell.style.fontWeight = 'bold';
        rowNumberCell.style.textAlign = 'center';
        tr.appendChild(rowNumberCell);

        row.forEach((cell, cellIndex) => {
          const td = document.createElement('td');
          td.contentEditable = "true";
          td.innerHTML = cell;
          td.dataset.row = rowIndex;
          td.dataset.cell = cellIndex;

          if (!cell || cell.trim() === '') {
            td.classList.add('empty-cell');
          }

          // Обработчик для клика - показ панели редактирования текста
          td.addEventListener('click', (e) => {
            e.stopPropagation();

            if (activeCell) {
              activeCell.classList.remove('active-edit');
            }

            activeCell = td;
            activeCell.classList.add('active-edit');
            activeTableIndex = idx;
            activeRowIndex = rowIndex;
            activeColumnIndex = cellIndex;

            toggleToolbar(true);
          });

          td.addEventListener('input', () => {
            if (td.innerHTML.trim() !== '') {
              td.classList.remove('empty-cell');
            } else {
              td.classList.add('empty-cell');
            }
            saveTableContent();
          });

          tr.appendChild(td);
        });
        editableTable.appendChild(tr);
      });
    }

    const wrapper = document.createElement('div');
    wrapper.className = "mb-8";
    wrapper.innerHTML = `
      <div class="table-header">
        <div class="flex items-center">
          <h2 class="text-xl font-semibold text-slate-800 table-title-editable table-title" 
              data-index="${idx}" contenteditable="true">${tableData.title}</h2>
          <span class="ml-2 text-sm text-gray-500">(ID: ${idx})</span>
        </div>
        <div class="table-actions">
          <button class="table-action-btn export-single-btn" data-index="${idx}">
            <i class="fas fa-file-export mr-1"></i> ${translations[currentLanguage].export}
          </button>
          <button class="table-action-btn column-btn" onclick="addColumn(${idx}, 'end')">
            <i class="fas fa-plus mr-1"></i> ${translations[currentLanguage].addColumn}
          </button>
          <button class="table-action-btn row-btn" onclick="addRow(${idx}, 'end')">
            <i class="fas fa-plus mr-1"></i> ${translations[currentLanguage].addRow}
          </button>
          <button class="table-action-btn delete-btn" onclick="deleteTable(${idx})">
            <i class="fas fa-trash mr-1"></i> ${translations[currentLanguage].deleteTable}
          </button>
        </div>
      </div>
      <div class="table-controls-row">
        <span class="table-controls-header">${translations[currentLanguage].columns}</span>
        <div class="column-controls">
          <button class="column-control-btn" onclick="addColumn(${idx}, 'start')">
            <i class="fas fa-arrow-left mr-1"></i>${translations[currentLanguage].addStart}
          </button>
          <button class="column-control-btn" onclick="addColumn(${idx}, 'end')">
            <i class="fas fa-arrow-right mr-1"></i>${translations[currentLanguage].addEnd}
          </button>
          <button class="column-control-btn" onclick="removeEmptyColumns(${idx})">
            <i class="fas fa-minus mr-1"></i>${translations[currentLanguage].empty}
          </button>
          <button class="delete-control-btn" onclick="removeCurrentColumn(${idx}, activeColumnIndex)" title="${translations[currentLanguage].deleteColumn}">
            <i class="fas fa-trash mr-1"></i>${translations[currentLanguage].current}
          </button>
        </div>
        
        <span class="table-controls-header">${translations[currentLanguage].rows}</span>
        <div class="row-controls">
          <button class="row-control-btn" onclick="addRow(${idx}, 'start')">
            <i class="fas fa-arrow-up mr-1"></i>${translations[currentLanguage].addStart}
          </button>
          <button class="row-control-btn" onclick="addRow(${idx}, 'end')">
            <i class="fas fa-arrow-down mr-1"></i>${translations[currentLanguage].addEnd}
          </button>
          <button class="row-control-btn" onclick="removeEmptyRows(${idx})">
            <i class="fas fa-minus mr-1"></i>${translations[currentLanguage].empty}
          </button>
          <button class="delete-control-btn" onclick="removeCurrentRow(${idx}, activeRowIndex)" title="${translations[currentLanguage].deleteRow}">
            <i class="fas fa-trash mr-1"></i>${translations[currentLanguage].current}
          </button>
        </div>
      </div>
    `;
    wrapper.appendChild(editableTable);
    tablesContainer.appendChild(wrapper);

    const titleElement = wrapper.querySelector('.table-title-editable');
    titleElement.addEventListener('blur', function () {
      updateTableTitle(idx, this.textContent);
    });

    const exportButton = wrapper.querySelector('.export-single-btn');
    exportButton.addEventListener('click', function () {
      const tableIndex = parseInt(this.getAttribute('data-index'));
      saveTableContent();
      if (extractedData[tableIndex]) {
        exportToExcel([extractedData[tableIndex]]);
      }
    });
  });
}

// Показать/скрыть панель инструментов
function toggleToolbar(show) {
  if (show) {
    textToolbar.classList.add('visible');
    document.body.style.paddingTop = '120px';
  } else {
    textToolbar.classList.remove('visible');
    document.body.style.paddingTop = '0';

    if (activeCell) {
      activeCell.classList.remove('active-edit');
      activeCell = null;
    }
  }
}

// Сохраняем выделение
function saveSelection() {
  if (window.getSelection) {
    const sel = window.getSelection();
    if (sel.getRangeAt && sel.rangeCount) {
      return sel.getRangeAt(0);
    }
  }
  return null;
}

// Восстанавливаем выделение
function restoreSelection(range) {
  if (range) {
    if (window.getSelection) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }
}

// Функция для применения размера шрифта
function applyFontSize(size) {
  if (!activeCell) return;

  const selection = saveSelection();
  activeCell.focus();

  if (selection) {
    restoreSelection(selection);
  }

  document.execCommand("styleWithCSS", false, true);

  const span = document.createElement('span');
  span.style.fontSize = size;

  if (selection && !selection.collapsed) {
    const selectedContent = selection.toString();
    if (selectedContent) {
      span.textContent = selectedContent;
      document.execCommand('insertHTML', false, span.outerHTML);
    }
  } else {
    activeCell.style.fontSize = size;
  }

  saveTableContent();
}

// Функция для применения цвета текста
function applyTextColor(colorValue) {
  if (!activeCell) return;

  const selection = saveSelection();
  activeCell.focus();

  if (selection) {
    restoreSelection(selection);

    // Современный подход вместо execCommand
    if (!selection.collapsed) {
      // Создаем span с нужным цветом для выделенного текста
      const span = document.createElement('span');
      span.style.color = colorValue;

      // Сохраняем выделенный текст
      const selectedText = selection.toString();
      if (selectedText) {
        span.textContent = selectedText;

        // Удаляем выделенный текст и вставляем форматированный
        const range = selection.getRangeAt(0);
        range.deleteContents();
        range.insertNode(span);
      }
    } else {
      // Если нет выделения, применяем ко всей ячейке
      activeCell.style.color = colorValue;
    }
  } else {
    // Если нет выделения, применяем ко всей ячейке
    activeCell.style.color = colorValue;
  }

  saveTableContent();
}
// Обработчики для кнопок панели инструментов
textToolbar.querySelectorAll('button').forEach(button => {
  if (!button.classList.contains('close-toolbar') && button.id !== 'colorPickerBtn') {
    button.addEventListener('click', () => {
      if (!activeCell) return;

      const command = button.dataset.command;
      const value = button.dataset.value;

      const selection = saveSelection();
      activeCell.focus();

      if (selection) {
        restoreSelection(selection);
      }

      if (value) {
        document.execCommand(command, false, value);
      } else {
        document.execCommand(command, false, null);
      }

      saveTableContent();
    });
  }
});

// Обработчики для выпадающих списков
fontNameSelect.addEventListener('change', () => {
  if (!activeCell) return;

  const selection = saveSelection();
  activeCell.focus();

  if (selection) {
    restoreSelection(selection);
  }

  document.execCommand('fontName', false, fontNameSelect.value);
  saveTableContent();
});

fontSizeSelect.addEventListener('change', () => {
  if (!activeCell) return;

  const size = fontSizeSelect.value;
  applyFontSize(size);
});

// Палитра цветов
colorPickerBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  colorPalette.classList.toggle('visible');
  isSelectingColor = true;
});

colorPalette.querySelectorAll('.palette-color').forEach(color => {
  color.addEventListener('click', (e) => {
    e.preventDefault();
    if (!activeCell) return;

    const colorValue = color.dataset.value;
    applyTextColor(colorValue);

    colorPalette.classList.remove('visible');
    isSelectingColor = false;
  });
});

// Закрытие панели инструментов
closeToolbarBtn.addEventListener('click', () => {
  toggleToolbar(false);
});

// Сохранение содержимого таблицы
function saveTableContent() {
  const edited = collectEditedTables();
  extractedData = edited;
}

function findNearestTitle(tableEl) {
  let cur = tableEl.previousElementSibling;
  let steps = 0;
  while (cur && steps < 8) {
    const txt = (cur.textContent || '').trim();
    if (txt.length > 0 && !cur.matches('table')) {
      return txt;
    }
    cur = cur.previousElementSibling;
    steps++;
  }
  return '';
}

function parseTables(container) {
  const tables = container.querySelectorAll('table');
  extractedData = [];

  tables.forEach((table, idx) => {
    const rows = Array.from(table.querySelectorAll('tr')).map(row =>
      Array.from(row.querySelectorAll('td,th')).map(cell => cell.textContent.trim())
    );

    let title = findNearestTitle(table);
    if (!title) title = `${currentLanguage === 'ru' ? 'Таблица' : 'Cədvəl'} ${idx + 1}`;

    extractedData.push({ title, table: rows });
  });

  renderTables();

  if (extractedData.length > 0) {
    downloadBtn.disabled = false;
    downloadExcelBtn.disabled = false;
    uploadBtn.disabled = false;
  }
}

function collectEditedTables() {
  const edited = [];
  const tableWrappers = tablesContainer.querySelectorAll('div.mb-8');

  tableWrappers.forEach((wrapper, idx) => {
    const titleElement = wrapper.querySelector('.table-title');
    const title = titleElement ? titleElement.textContent : `${currentLanguage === 'ru' ? 'Таблица' : 'Cədvəl'} ${idx + 1}`;

    const rows = [];
    const tableRows = wrapper.querySelectorAll('tr');

    tableRows.forEach(row => {
      // Пропускаем первую ячейку с номером строки
      const cells = Array.from(row.querySelectorAll('td')).slice(1).map(cell => cell.innerHTML.trim());
      rows.push(cells);
    });

    edited.push({ title, table: rows });
  });

  return edited;
}

// Обработчики для модального окна экспорта
downloadExcelBtn.addEventListener('click', showExportModal);

selectAllCheckbox.addEventListener('change', function () {
  const checkboxes = tableList.querySelectorAll('.table-checkbox');
  checkboxes.forEach(checkbox => {
    checkbox.checked = this.checked;
  });
});

cancelExportBtn.addEventListener('click', function () {
  exportModal.classList.remove('visible');
});

confirmExportBtn.addEventListener('click', function () {
  const checkboxes = tableList.querySelectorAll('.table-checkbox:checked');
  selectedTablesForExport = Array.from(checkboxes).map(checkbox => {
    const index = parseInt(checkbox.getAttribute('data-index'));
    return extractedData[index];
  });

  exportModal.classList.remove('visible');
  exportToExcel(selectedTablesForExport);
});

input.addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async function (e) {
    const buffer = e.target.result;
    const result = await window.mammoth.convertToHtml({ arrayBuffer: buffer });
    const temp = document.createElement('div');
    temp.innerHTML = result.value;
    parseTables(temp);
  };
  reader.readAsArrayBuffer(file);
});

downloadBtn.addEventListener('click', () => {
  const edited = collectEditedTables();
  const blob = new Blob([JSON.stringify(edited, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'lesson_plan.json';
  a.click();
  URL.revokeObjectURL(url);
});

uploadBtn.addEventListener('click', async () => {
  const edited = collectEditedTables();
  try {
    await fetch('/upload', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(edited)
    });
    alert(currentLanguage === 'ru' ? 'Таблица отправлена на сервер!' : 'Cədvəl serverə göndərildi!');
  } catch (err) {
    alert((currentLanguage === 'ru' ? 'Ошибка отправки: ' : 'Göndərilmə xətası: ') + err.message);
  }
});

// Скрываем панель инструментов при клике вне её области
document.addEventListener('click', (e) => {
  if (!textToolbar.contains(e.target) && !e.target.closest('td[contenteditable]')) {
    toggleToolbar(false);
  }

  if (!colorPalette.contains(e.target) && e.target !== colorPickerBtn) {
    colorPalette.classList.remove('visible');
    isSelectingColor = false;
  }

  if (exportModal.classList.contains('visible') && e.target === exportModal) {
    exportModal.classList.remove('visible');
  }
  
  if (instructionModal.classList.contains('visible') && e.target === instructionModal) {
    instructionModal.classList.remove('visible');
  }
});

// Глобальные функции для вызова из onclick
window.addColumn = addColumn;
window.removeEmptyColumns = removeEmptyColumns;
window.removeCurrentColumn = removeCurrentColumn;
window.addRow = addRow;
window.removeEmptyRows = removeEmptyRows;
window.removeCurrentRow = removeCurrentRow;
window.deleteTable = deleteTable;

// Инициализация языка
setLanguage('ru');
