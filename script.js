"use strict";
let table = document.getElementsByClassName("sheet-body")[0];
let rows = document.querySelector(".rows");
let columns = document.querySelector(".columns");
let generateBtn = document.querySelector(".generate-btn");
let newSheetBtn = document.querySelector(".new-sheet-btn");
let tableExists = false;

generateBtn.addEventListener("click", () => {
let rowsNumber = rows.value;
let columnsNumber = columns.value;
table.innerHTML = "";

let rowInvalid = rowsNumber === "";
let ColumnInvalid = columnsNumber === "";
  let message = `First You must insert number of `;
  
  if (rowInvalid || ColumnInvalid) {
    Swal.fire({
      title: "Sorry!",
      text: `${
        rowInvalid && ColumnInvalid
          ? `${message} { rows and columns }`
          : `${message}{ ${rowInvalid ? "rows" : ""} ${
              ColumnInvalid ? "colums" : ""
            }}`
      }`,
      icon: "warning",
      iconColor: "red",
      confirmButtonColor: "#bd05bd",
      confirmButtonText: "Back to Page",
      background: "#fae0fa",
      width: "25rem",
    }).then((result) => {
      if (result.isConfirmed) {
        rows.focus();
      }
    });
  }
  
  for (let i = 0; i < rowsNumber; i++) {
    let tableRow = "";
    for (let j = 0; j < columnsNumber; j++) {
      tableRow += `<td contenteditable></td>`;
    }
    table.innerHTML += tableRow;
  }
  if (rowsNumber > 0 && columnsNumber > 0) {
    tableExists = true;
    rows.value = "";
    columns.value = "";
    generateBtn.classList.add("hidden-btn");
    newSheetBtn.classList.remove("hidden-btn");
  }
})

const ExportToExcel = (type, fn, dl) => {
  if (!tableExists) {
    Swal.fire({
      title: "Ooops!",
      text: "First You must generate Excel Sheet",
      icon: "warning",
      iconColor: "red",
      confirmButtonColor: "#bd05bd",
      confirmButtonText: "Back to Page",
      background: "#fae0fa",
      width: "25rem",
    });
  }
  let elt = table;
  let wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
  return dl
    ? XLSX.write(wb, { bookType: type, bookSST: true, type: "base64" })
    : XLSX.writeFile(wb, fn || "MyNewSheet." + (type || "xlsx"));
};

function newSheet() {
  rows.value = "";
  columns.value = "";
  table.innerHTML = "";
  tableExists = false;
  newSheetBtn.classList.add("hidden-btn");
   generateBtn.classList.remove("hidden-btn");
}
