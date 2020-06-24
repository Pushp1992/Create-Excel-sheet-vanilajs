/**
 * Global Variable
 */
let rCount = localStorage.getItem('rowCount');
let cCount = localStorage.getItem('colCount');

/**
 * Showing/hiding clear button
 */
if (localStorage.getItem('rowCount') !== null) {
    document.getElementById('clear-btn').style.visibility = "visible";
    document.getElementById('addRowBtn').style.visibility = "visible";
    document.getElementById('removeRowBtn').style.visibility = "visible";
    document.getElementById('addColBtn').style.visibility = "visible";
    document.getElementById('removeColBtn').style.visibility = "visible";
} else {
    document.getElementById('clear-btn').style.visibility = "hidden";
    document.getElementById('addRowBtn').style.visibility = "hidden";
    document.getElementById('removeRowBtn').style.visibility = "hidden";
    document.getElementById('addColBtn').style.visibility = "hidden";
    document.getElementById('removeColBtn').style.visibility = "hidden";
}

/**
 * Clear excel sheet
 */
const clearExcelFile = function (event) {
    event.preventDefault();
    localStorage.clear()
    location.reload();
}

/**
 * Generate Excel Sheet based on Row and column input Count
 */

const generateExcel = function (event) {
    event.preventDefault();

    let existingRowCount = localStorage.getItem('rowCount');
    let existingColCount = localStorage.getItem('colCount');

    console.log(existingRowCount, existingColCount)

    let rowCountFromUser = document.getElementById("rowCount").value;
    let columnCountFromUser = document.getElementById("colCount").value;

    let rowCount = rowCountFromUser === "" ? (existingRowCount !== null ? existingRowCount : "") : rowCountFromUser;
    let colCount = columnCountFromUser === "" ? (existingColCount !== null ? existingColCount : "") : columnCountFromUser;

    if (rowCount === "" || rowCount <= 0) {
        window.alert("Row count cannnot be empty or less than 0");
        return false;
    }
    if (colCount === "" || colCount <= 0) {
        window.alert("Column count cannnot be empty or less than 0");
        return false;
    }

    localStorage.setItem('rowCount', rowCount);
    localStorage.setItem('colCount', colCount);
    location.reload();
}

/**
 * Logic for adding/removing row explicitly
 */
const addRow = function(event) {
    event.preventDefault();

    let incrementedRow = parseInt(rCount) + 1 ;
    localStorage.setItem('rowCount', incrementedRow);
    location.reload();
}

const removeRow = function(event) {
    event.preventDefault();

    let decrementedRow = parseInt(rCount) - 1 ;
    localStorage.setItem('rowCount', decrementedRow);
    location.reload();
}

/**
 * Adding removing column explicitly
 */
const addColumn = function(event) {
    event.preventDefault();

    let incrementedColumn = parseInt(cCount) + 1 ;
    localStorage.setItem('colCount', incrementedColumn);
    location.reload();
}

const removeColumn = function(event) {
    event.preventDefault();

    let decrementedColumn = parseInt(cCount) - 1 ;
    localStorage.setItem('colCount', decrementedColumn);
    location.reload();
}

/**
 * Logic for editing cell and stroing data in excel sheet
 */
// let rowCounter = 0;
// let colCounter = 0;
// let rCount = localStorage.getItem('rowCount');
// let cCount = localStorage.getItem('colCount');

for (var i = 0; i <= rCount; i++) {
    var row = document.querySelector("table").insertRow(-1);
    for (var j = 0; j <= cCount; j++) {
        var letter = String.fromCharCode("A".charCodeAt(0) + j - 1);
        row.insertCell(-1).innerHTML = i && j ? "<input class='cellInputField' id='" + letter + i + "'/> " : i || letter;
    }
}

var DATA = {}, excelInputFields = [].slice.call(document.getElementsByClassName("cellInputField"));
excelInputFields.forEach(function (cell) {
    cell.onfocus = function (e) {
        e.target.value = localStorage[e.target.id] || "";
    };
    cell.onblur = function (e) {
        localStorage[e.target.id] = e.target.value;
        computeAll();
    };
    var getCellData = function () {
        var value = localStorage[cell.id] || "";
        if (value.charAt(0) == "=") {
            with (DATA) return eval(value.substring(1));
        } else { return isNaN(parseFloat(value)) ? value : parseFloat(value); }
    };
    Object.defineProperty(DATA, cell.id, { get: getCellData });
    Object.defineProperty(DATA, cell.id.toLowerCase(), { get: getCellData });
});
(window.computeAll = function () {
    excelInputFields.forEach(function (cell) { try { cell.value = DATA[cell.id]; } catch (e) { } });
})();