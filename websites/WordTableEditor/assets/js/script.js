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

    let extractedData = [];
    let activeCell = null;
    let isSelectingColor = false;
    let selectedTablesForExport = [];

    // Функция для экспорта в Excel с выбранными таблицами
    function exportToExcel(tablesToExport = null) {
      if (extractedData.length === 0) {
        alert('Нет данных для экспорта!');
        return;
      }

      // Если не указаны конкретные таблицы, экспортируем все
      const tables = tablesToExport || extractedData;
      
      if (tables.length === 0) {
        alert('Не выбрано ни одной таблицы для экспорта!');
        return;
      }

      try {
        // Создаем новую рабочую книгу
        const wb = XLSX.utils.book_new();
        
        // Для каждой таблицы создаем отдельный лист
        tables.forEach((tableData, index) => {
          // Подготавливаем данные для Excel
          const excelData = [];
          
          // Добавляем заголовок таблицы как первую строку
          excelData.push([tableData.title]);
          excelData.push([]); // Пустая строка для разделения
          
          // Добавляем данные таблицы
          if (tableData.table && tableData.table.length > 0) {
            // Добавляем строки с данными
            tableData.table.forEach(row => {
              // Убираем HTML теги из текста для чистого Excel
              const cleanRow = row.map(cell => {
                const div = document.createElement('div');
                div.innerHTML = cell;
                return div.textContent || div.innerText || '';
              });
              excelData.push(cleanRow);
            });
          }
          
          // Создаем worksheet
          const ws = XLSX.utils.aoa_to_sheet(excelData);
          
          // Добавляем стили для заголовка (первая строка)
          if (!ws['A1']) ws['A1'] = {t: 's'};
          ws['A1'].s = {
            font: {bold: true, sz: 14},
            alignment: {horizontal: 'center'}
          };
          
          // Устанавливаем ширину столбцов
          const colWidths = [];
          if (tableData.table && tableData.table.length > 0) {
            const firstDataRow = tableData.table[0];
            for (let i = 0; i < firstDataRow.length; i++) {
              colWidths.push({wch: 20}); // Фиксированная ширина колонок
            }
            ws['!cols'] = colWidths;
          }
          
          // Добавляем границы для всех ячеек с данными
          const range = XLSX.utils.decode_range(ws['!ref']);
          for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
              const cell_address = {c: C, r: R};
              const cell_ref = XLSX.utils.encode_cell(cell_address);
              
              if (!ws[cell_ref]) continue;
              
              if (!ws[cell_ref].s) ws[cell_ref].s = {};
              ws[cell_ref].s.border = {
                top: {style: 'thin'},
                bottom: {style: 'thin'},
                left: {style: 'thin'},
                right: {style: 'thin'}
              };
              
              // Жирный шрифт для заголовков таблиц (первая строка данных)
              if (R === 2 && tableData.table && tableData.table[0]) {
                if (!ws[cell_ref].s.font) ws[cell_ref].s.font = {};
                ws[cell_ref].s.font.bold = true;
                ws[cell_ref].s.fill = {fgColor: {rgb: "E6E6FA"}}; // Светло-фиолетовый фон
              }
            }
          }
          
          // Добавляем worksheet в книгу
          let sheetName = tableData.title.replace(/[\\/*?:[\]]/g, '').substring(0, 31);
          if (sheetName === '') sheetName = `Таблица ${index + 1}`;
          
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });
        
        // Сохраняем файл
        const fileName = `lesson_plan_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        
      } catch (error) {
        console.error('Ошибка при экспорте в Excel:', error);
        alert('Произошла ошибка при создании Excel файла: ' + error.message);
      }
    }

    // Показать модальное окно выбора таблиц
    function showExportModal() {
      // Очищаем список таблиц
      tableList.innerHTML = '';
      
      // Заполняем список таблиц с уникальными идентификаторами
      extractedData.forEach((table, index) => {
        const listItem = document.createElement('div');
        listItem.className = 'table-list-item';
        listItem.innerHTML = `
          <input type="checkbox" class="table-checkbox" id="modal-table-${index}" data-index="${index}" checked>
          <label for="modal-table-${index}" class="flex-1">${table.title} (ID: ${index})</label>
        `;
        tableList.appendChild(listItem);
      });
      
      // Устанавливаем checkbox "Выбрать все" в отмеченное состояние
      selectAllCheckbox.checked = true;
      
      // Показываем модальное окно
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
      if (confirm('Вы уверены, что хотите удалить эту таблицу?')) {
        // Удаляем из данных
        extractedData.splice(index, 1);
        
        // Перерисовываем интерфейс
        renderTables();
        
        // Сохраняем изменения
        saveTableContent();
        
        // Если таблиц не осталось, деактивируем кнопки
        if (extractedData.length === 0) {
          downloadBtn.disabled = true;
          downloadExcelBtn.disabled = true;
          uploadBtn.disabled = true;
        }
      }
    }

    // Перерисовка таблиц в интерфейсе
    function renderTables() {
      tablesContainer.innerHTML = '';
      
      extractedData.forEach((tableData, idx) => {
        // Создаём редактируемую таблицу
        const editableTable = document.createElement('table');
        if (tableData.table && tableData.table.length > 0) {
          tableData.table.forEach(row => {
            const tr = document.createElement('tr');
            row.forEach(cell => {
              const td = document.createElement('td');
              td.contentEditable = "true";
              td.innerHTML = cell;
              tr.appendChild(td);

              // Добавляем обработчик для показа панели инструментов
              td.addEventListener('click', (e) => {
                // Убираем активный класс у предыдущей ячейки
                if (activeCell) {
                  activeCell.classList.remove('active-edit');
                }
                
                // Устанавливаем новую активную ячейку
                activeCell = td;
                activeCell.classList.add('active-edit');
                
                // Показываем панель инструментов вверху экрана
                toggleToolbar(true);
              });
              
              // Добавляем обработчик ввода текста
              td.addEventListener('input', () => {
                saveTableContent();
              });
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
                <i class="fas fa-file-export mr-1"></i> Экспорт
              </button>
              <button class="table-action-btn delete-btn" data-index="${idx}">
                <i class="fas fa-trash mr-1"></i> Удалить
              </button>
            </div>
          </div>
        `;
        wrapper.appendChild(editableTable);
        tablesContainer.appendChild(wrapper);

        // Добавляем обработчик для редактирования заголовка
        const titleElement = wrapper.querySelector('.table-title-editable');
        titleElement.addEventListener('blur', function() {
          updateTableTitle(idx, this.textContent);
        });

        // Добавляем обработчик для экспорта одной таблицы
        const exportButton = wrapper.querySelector('.export-single-btn');
        exportButton.addEventListener('click', function() {
          const tableIndex = parseInt(this.getAttribute('data-index'));
          // Сначала сохраняем текущие данные
          saveTableContent();
          if (extractedData[tableIndex]) {
            exportToExcel([extractedData[tableIndex]]);
          }
        });
        
        // Добавляем обработчик для удаления таблицы
        const deleteButton = wrapper.querySelector('.delete-btn');
        deleteButton.addEventListener('click', function() {
          const tableIndex = parseInt(this.getAttribute('data-index'));
          deleteTable(tableIndex);
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
      
      // Сохраняем выделение
      const selection = saveSelection();
      
      activeCell.focus();
      
      // Восстанавливаем выделение
      if (selection) {
        restoreSelection(selection);
      }
      
      // Применяем размер через CSS
      document.execCommand("styleWithCSS", false, true);
      
      // Создаем span с нужным размером шрифта
      const span = document.createElement('span');
      span.style.fontSize = size;
      
      // Если есть выделение, обрабатываем его
      if (selection && !selection.collapsed) {
        const selectedContent = selection.toString();
        if (selectedContent) {
          span.textContent = selectedContent;
          document.execCommand('insertHTML', false, span.outerHTML);
        }
      } else {
        // Если нет выделения, применяем ко всей ячейке
        activeCell.style.fontSize = size;
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
          
          // Сохраняем выделение
          const selection = saveSelection();
          
          // Фокусируемся на активной ячейке
          activeCell.focus();
          
          // Восстанавливаем выделение
          if (selection) {
            restoreSelection(selection);
          }
          
          // Выполняем команду документа
          if (value) {
            document.execCommand(command, false, value);
          } else {
            document.execCommand(command, false, null);
          }
          
          // Сохраняем изменения
          saveTableContent();
        });
      }
    });

    // Обработчики для выпадающих списков
    fontNameSelect.addEventListener('change', () => {
      if (!activeCell) return;
      
      // Сохраняем выделение
      const selection = saveSelection();
      
      activeCell.focus();
      
      // Восстанавливаем выделение
      if (selection) {
        restoreSelection(selection);
      }
      
      document.execCommand('fontName', false, fontNameSelect.value);
      saveTableContent();
    });

    fontSizeSelect.addEventListener('change', () => {
      if (!activeCell) return;
      
      const size = fontSizeSelect.value;
      
      // Применяем размер шрифта
      applyFontSize(size);
    });

    // Палитра цветов
    colorPickerBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      colorPalette.classList.toggle('visible');
      isSelectingColor = true;
    });

    // Обработчики для выбора цвета из палитры
    colorPalette.querySelectorAll('.palette-color').forEach(color => {
      color.addEventListener('click', (e) => {
        if (!activeCell) return;
        
        const colorValue = color.dataset.value;
        
        // Сохраняем выделение
        const selection = saveSelection();
        
        activeCell.focus();
        
        // Восстанавливаем выделение
        if (selection) {
          restoreSelection(selection);
        }
        
        document.execCommand('foreColor', false, colorValue);
        saveTableContent();
        
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
        if (!title) title = `Таблица ${idx + 1}`;

        extractedData.push({ title, table: rows });
      });

      // Рендерим таблицы
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
        const title = titleElement ? titleElement.textContent : `Таблица ${idx + 1}`;
        
        const rows = [];
        const tableRows = wrapper.querySelectorAll('tr');
        
        tableRows.forEach(row => {
          const cells = Array.from(row.querySelectorAll('td')).map(cell => cell.innerHTML.trim());
          rows.push(cells);
        });
        
        edited.push({ title, table: rows });
      });
      
      return edited;
    }

    // Обработчики для модального окна экспорта
    downloadExcelBtn.addEventListener('click', showExportModal);

    selectAllCheckbox.addEventListener('change', function() {
      const checkboxes = tableList.querySelectorAll('.table-checkbox');
      checkboxes.forEach(checkbox => {
        checkbox.checked = this.checked;
      });
    });

    cancelExportBtn.addEventListener('click', function() {
      exportModal.classList.remove('visible');
    });

    confirmExportBtn.addEventListener('click', function() {
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
      reader.onload = async function(e) {
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
        alert('Таблица отправлена на сервер!');
      } catch (err) {
        alert('Ошибка отправки: ' + err.message);
      }
    });

    // Скрываем панель инструментов при клике вне её области
    document.addEventListener('click', (e) => {
      if (!textToolbar.contains(e.target) && !e.target.closest('td[contenteditable]')) {
        toggleToolbar(false);
      }
      
      // Скрываем палитру цветов при клике вне её
      if (!colorPalette.contains(e.target) && e.target !== colorPickerBtn) {
        colorPalette.classList.remove('visible');
        isSelectingColor = false;
      }
      
      // Скрываем модальное окно при клике вне его
      if (exportModal.classList.contains('visible') && e.target === exportModal) {
        exportModal.classList.remove('visible');
      }
    });