function chooseFile() {
    document.getElementById('fileInput').click();
}

function handleFile(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, { type: 'binary' });
        var sheetNames = workbook.SheetNames; // Отримати список імен аркушів у книзі
        var sheetName = sheetNames[0]; // Взяти перше вкладення
        var sheet = workbook.Sheets[sheetName]; // Взяти його ім'я

        let lastRow = findLastNonEmptyCell(sheet).row;
        let lastColumn = XLSX.utils.encode_col(findLastNonEmptyCell(sheet).column);

        var range = `A1:${lastColumn}${lastRow}`; // Визначте діапазон комірок
        var selectedRange = XLSX.utils.sheet_to_json(sheet, { range: range, header: 1 }); // Отримайте значення з визначеного діапазону

        var result = getSum(selectedRange); // Сума
        let stakes = stakesByComa(selectedRange); // Ставки через кому

        // Вивести значення в елемент з ідентифікатором "output-sum"
        //document.querySelector('.output-sum').innerHTML = 'Повернуто: ' + result.sumReturn + '<br> Збережено: ' + result.sumSaved;
        document.querySelector('.output-sumReturn').innerHTML = result.sumReturn;
        document.querySelector('.output-sumSaved').innerHTML = result.sumSaved;
        document.querySelector('.output-stakes').innerHTML = stakes.stakes;
        document.querySelector('.output-stakes-count').innerHTML = stakes.stakesCount;

        console.log(stakesByComa(selectedRange));
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

function findLastNonEmptyCell(sheet) {
    var range = XLSX.utils.decode_range(sheet['!ref']);

    for (var rowNum = range.e.r; rowNum >= range.s.r; rowNum--) {
        for (var colNum = range.e.c; colNum >= range.s.c; colNum--) {
            var cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
            var cell = sheet[cellAddress];

            if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
                // Знайдено непорожню комірку, повертаємо об'єкт з номером рядка і колонки
                return { row: rowNum + 1, column: colNum + 1 };
            }
        }
    }

    // Якщо всі комірки порожні, повертаємо null або інше значення за замовчуванням
    return null;
}

function getSum(array) {
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

    // Повернути значення як об'єкт
    return {
        'sumReturn': sumReturn.toFixed(2),
        'sumSaved': sumSaved.toFixed(2)
    };
}

function stakesByComa(array) {
    let idIndex = array[0].indexOf('Id');
    let stakes = [];

    for (let i = 1; i < array.length; i++) {
        stakes.push(array[i][idIndex]);
    }

    return {'stakes': stakes.join(', '), 'stakesCount': stakes.length};
}

document.querySelector('.clearBtn').addEventListener('click', clear)

function clear() {
    document.getElementById('fileInput').value = '';

    document.querySelector('.output-sumReturn').innerHTML = '';
    document.querySelector('.output-sumSaved').innerHTML = '';
    document.querySelector('.output-stakes').innerHTML = '';
    document.querySelector('.output-stakes-count').innerHTML = '';
}