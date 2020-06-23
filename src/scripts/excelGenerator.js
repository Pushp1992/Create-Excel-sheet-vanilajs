/**
 * Showing/hiding clear button
 */
if (localStorage.getItem('rowCount') !== null) {
    document.getElementById('clear-btn').style.visibility = "visible";
} else {
    document.getElementById('clear-btn').style.visibility = "hidden";
}

/**
 * Clear excel sheet
 */
const clearExcelFile = function (event) {
    event.preventDefault();
    localStorage.clear()
    location.reload();
}

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
 * Generate Excel Sheet based on Row and column input Count
 */

let rCount = localStorage.getItem('rowCount');
let cCount = localStorage.getItem('colCount');

for (var i = 0; i <= rCount; i++) {
    var row = document.querySelector("table").insertRow(-1);

    // var cell1 = row.insertCell(-1);
    // var button = document.createElement("input");
    // button.setAttribute("type", "button")
    // button.setAttribute("name", "Delete");
    // button.setAttribute("value", "del");
    // button.setAttribute("onclick", `removeRow()`);
    // cell1.appendChild(button);

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