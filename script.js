function chooseFile() {
            document.getElementById('fileInput').click();
        }

        function handleFile(file) {
            var reader = new FileReader();
            reader.onload = function (e) {
                var data = e.target.result;
                var workbook = XLSX.read(data, { type: 'binary' });

                // Отримати список імен аркушів у книзі
                var sheetNames = workbook.SheetNames;

                // Вивести імена аркушів в консоль
                console.log('Імена аркушів:', sheetNames);

                // Визначте ім'я аркуша, з якого ви хочете отримати дані
                var sheetName = sheetNames[0];
                var sheet = workbook.Sheets[sheetName];

                // Знайти останній не порожній рядок в аркуші
                var lastNonEmptyRow = findLastNonEmptyRow(sheet);
                console.log('Останній не порожній рядок:', lastNonEmptyRow);

                // Визначте номер рядка, в якому ви хочете знайти останню непорожню колонку
                var rowNumber = 1; // Наприклад, рядок 1

                var lastNonEmptyColumn = findLastNonEmptyColumn(sheet, rowNumber);

                if (lastNonEmptyColumn !== null) {
                    console.log('Остання непорожня колонка:', XLSX.utils.encode_col(lastNonEmptyColumn));
                } else {
                    console.log('У вказаному рядку всі колонки порожні.');
                }

                console.log(lastNonEmptyRow, XLSX.utils.encode_col(lastNonEmptyColumn));

                // Визначте діапазон комірок, наприклад, 'A1:B5'
                var range = 'A2:B6';
                var range_2 = `A1:${XLSX.utils.encode_col(lastNonEmptyColumn)}${lastNonEmptyRow}`;

                // Отримайте значення з визначеного діапазону
                var selectedRange = XLSX.utils.sheet_to_json(sheet, { range: range, header: 1 });
                var selectedRange_2 = XLSX.utils.sheet_to_json(sheet, { range: range_2, header: 1 });

                console.log('Значення обраного діапазону:', selectedRange);
                console.log('Значення обраного діапазону:', selectedRange_2);

                // Сума
                var result = sum(selectedRange_2);

                // Вивести значення в елемент з ідентифікатором "output"
                document.getElementById('output').innerHTML = 'sumReturn: ' + result.sumReturn + '<br> sumSaved: ' + result.sumSaved;
            };

            reader.readAsBinaryString(file);
        }

        function handleDrop(event) {
            event.preventDefault();
            var files = event.dataTransfer.files;
            handleFile(files[0]);
        }

        function handleInputChange(event) {
            var files = event.target.files;
            handleFile(files[0]);
        }

        var dropArea = document.getElementById('dropArea');

        dropArea.addEventListener('dragover', function (event) {
            event.preventDefault();
            dropArea.style.border = '2px dashed #aaa';
        });

        dropArea.addEventListener('dragleave', function () {
            dropArea.style.border = '2px dashed #ccc';
        });

        dropArea.addEventListener('drop', handleDrop);

        document.getElementById('fileInput').addEventListener('change', handleInputChange);

        // Пошук останнього рядка
        function findLastNonEmptyRow(sheet) {
            var range = XLSX.utils.decode_range(sheet['!ref']);

            for (var rowNum = range.e.r; rowNum >= range.s.r; rowNum--) {
                var row = [];
                for (var colNum = range.s.c; colNum <= range.e.c; colNum++) {
                    var cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
                    var cell = sheet[cellAddress];
                    if (cell && cell.v !== undefined) {
                        // Знайдено непорожню комірку, це останній не порожній рядок
                        return rowNum + 1;
                    }
                }
            }

            // Якщо рядки повністю порожні, повернути 0 або інший індикатор, залежно від потреб
            return 0;
        }

        // Пошук останньої колонки
        function findLastNonEmptyColumn(sheet, rowNumber) {
            var range = XLSX.utils.decode_range(sheet['!ref']);

            for (var col = range.e.c; col >= range.s.c; col--) {
                var cellAddress = { r: rowNumber, c: col };
                var cellRef = XLSX.utils.encode_cell(cellAddress);
                var cell = sheet[cellRef];

                // Перевірка, чи комірка не порожня
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    return col;
                }
            }
            // Якщо всі комірки порожні, повертаємо null або інше значення за замовчуванням
            return null;
        }

        function sum(array) {
            let coefIndex = array[0].indexOf('Coef');
            let sumInIndex = array[0].indexOf('Sum in');

            let sumReturn = 0;
            let sumSaved = 0;

            for (let i = 1; i < array.length; i++) {
                let sumItem = '' + array[i][sumInIndex];
                let coefItem = '' + array[i][coefIndex];

                sumReturn += sumItem * 1;
                sumSaved += (sumItem * 1) * ((coefItem * 1) - 1);
            }

            // Вивести значення на консоль
            console.log('sumReturn:', sumReturn);
            console.log('sumSaved:', sumSaved);

            // Вивести значення в елемент з ідентифікатором "output"
            document.getElementById('output').innerHTML = 'sumReturn: ' + sumReturn + '<br> sumSaved: ' + sumSaved;

            // Повернути значення як об'єкт
            return {
                'sumReturn': sumReturn,
                'sumSaved': sumSaved
            };
        }