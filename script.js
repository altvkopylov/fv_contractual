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

        let stakes = stakesByComa(selectedRange); // Ставки через кому

        console.log(groupByCurrency(selectedRange))

        console.log(groupByCurrency(selectedRange));
        saveToLocalStorage(selectedRange);

        let groupedData = groupByCurrency(selectedRange).groupedData;
        document.querySelector('.output-return-saved').innerHTML = formatResult(groupedData);

        document.querySelector('.output-stakes').innerHTML = (stakes.stakes) ? stakes.stakes : '';
        document.querySelector('.output-stakes-count').innerHTML = (stakes.stakesCount) ? stakes.stakesCount : '';

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

    return null;
}

function stakesByComa(array) {
    let idIndex = array[0].indexOf('Id');
    
    if (idIndex == -1) {
        return document.querySelector('.error').innerHTML = 'Неправильний тип файлу. Немає колонки "Id"';
    }

    let stakes = [];

    for (let i = 1; i < array.length; i++) {
        stakes.push(array[i][idIndex]);
    }

    return {'stakes': stakes.join(', '), 'stakesCount': stakes.length};
}

document.querySelector('.clearBtn').addEventListener('click', clear)

function clear() {
    document.getElementById('fileInput').value = '';

    document.querySelector('.output-return-saved').innerHTML = '';
    document.querySelector('.output-stakes').innerHTML = '';
    document.querySelector('.output-stakes-count').innerHTML = '';

    document.querySelector('.error').innerHTML = '';

    localStorage.clear();
}

function saveToLocalStorage(array) {

    (groupByCurrency(array)) ? localStorage.setItem('result', JSON.stringify(groupByCurrency(array).groupedData)) : '';
    (stakesByComa(array).stakes) ? localStorage.setItem('stakes', stakesByComa(array).stakes) : '';
    (stakesByComa(array).stakesCount) ? localStorage.setItem('stakesCount', stakesByComa(array).stakesCount) : '';
}

function getByLocalStorage() {
    resultData = localStorage.getItem('result');
    document.querySelector('.output-return-saved').innerHTML = formatResult(JSON.parse(resultData))
    document.querySelector('.output-stakes').innerHTML = localStorage.getItem('stakes');
    document.querySelector('.output-stakes-count').innerHTML = localStorage.getItem('stakesCount');
}

document.addEventListener('DOMContentLoaded', getByLocalStorage)



function groupByCurrency(data) {
    let coefIndex = data[0].indexOf('Coef');
    let sumInIndex = data[0].indexOf('Sum in');
    let CurrencyIndex = data[0].indexOf('Currency');

    if (coefIndex == -1 || sumInIndex == -1 || CurrencyIndex == -1) {
        return document.querySelector('.error').innerHTML = 'Неправильний тип файлу';
    }

    let groupedData = {};
    for (let i = 1; i < data.length; i++) {
        let row = data[i];
        let currency = row[data[0].indexOf('Currency')];
        if (!groupedData[currency]) {
            groupedData[currency] = {
                currency: currency,
                sumReturn: 0,
                sumSaved: 0
            };
        }
        groupedData[currency].sumReturn += Math.round(row[data[0].indexOf('Sum in')]);
        groupedData[currency].sumSaved += Math.round(row[data[0].indexOf('Sum in')] * (row[data[0].indexOf('Coef')] - 1));
    }
    let result = Object.values(groupedData);
    return {'groupedData': result};
}

function formatResult(data) {
    let resultString = '';
    for (let currency in data) {
        resultString += `<span class="currency">${data[currency].currency}</span>` + ':<br>';
        resultString += 'Повернуто: ' + data[currency].sumReturn + '<br>';
        resultString += 'Збережено: ' + data[currency].sumSaved + '<br><br>';
    }
    return resultString;
}