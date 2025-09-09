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

    let extractedData = [];
    let activeCell = null;
    let isSelectingColor = false;

    // Функция для экспорта в Excel
    function exportToExcel() {
      if (extractedData.length === 0) {
        alert('Нет данных для экспорта!');
        return;
      }

      try {
        // Создаем новую рабочую книгу
        const wb = XLSX.utils.book_new();
        
        // Для каждой таблицы создаем отдельный лист
        extractedData.forEach((tableData, index) => {
          // Подготавливаем данные для Excel
          const excelData = [];
          
          // Добавляем заголовок таблицы как первую строку
          excelData.push([tableData.title]);
          excelData.push([]); // Пустая строка для разделения
          
          // Добавляем данные таблицы
          tableData.table.forEach(row => {
            excelData.push(row);
          });
          
          // Создаем worksheet
          const ws = XLSX.utils.aoa_to_sheet(excelData);
          
          // Добавляем worksheet в книгу
          // Используем безопасное имя для листа (максимум 31 символ, без запрещенных символов)
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
      tablesContainer.innerHTML = '';

      tables.forEach((table, idx) => {
        const rows = Array.from(table.querySelectorAll('tr')).map(row =>
          Array.from(row.querySelectorAll('td,th')).map(cell => cell.textContent.trim())
        );

        let title = findNearestTitle(table);
        if (!title) title = `Таблица ${idx + 1}`;

        extractedData.push({ title, table: rows });

        // Создаём редактируемую таблицу
        const editableTable = document.createElement('table');
        rows.forEach(row => {
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

        const wrapper = document.createElement('div');
        wrapper.className = "mb-8";
        wrapper.innerHTML = `<h2 class="text-xl font-semibold mb-3 text-slate-800">${title}</h2>`;
        wrapper.appendChild(editableTable);
        tablesContainer.appendChild(wrapper);
      });

      if (extractedData.length > 0) {
        downloadBtn.disabled = false;
        downloadExcelBtn.disabled = false;
        uploadBtn.disabled = false;
      }
    }

    function collectEditedTables() {
      const edited = [];
      tablesContainer.querySelectorAll('div').forEach((wrapper, idx) => {
        const title = wrapper.querySelector('h2')?.textContent || `Таблица ${idx + 1}`;
        const rows = Array.from(wrapper.querySelectorAll('tr')).map(row =>
          Array.from(row.querySelectorAll('td')).map(cell => cell.innerHTML.trim())
        );
        edited.push({ title, table: rows });
      });
      return edited;
    }

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

    // Обработчик для кнопки скачивания Excel
    downloadExcelBtn.addEventListener('click', exportToExcel);

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
    });