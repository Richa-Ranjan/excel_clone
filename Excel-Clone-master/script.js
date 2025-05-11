let defaultProperties = {
    text: "",
    "font-weight": "",
    "font-style": "",
    "text-decoration": "",
    "text-align": "left",
    "background-color": "#ffffff",
    "color": "#000000",
    "font-family": "Noto Sans",
    "font-size": "14px"
}

let cellData = {
    "Sheet1": {}
}

let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastlyAddedSheet = 1;
$(document).ready(function () {
    let cellContainer = $(".input-cell-container");

    for (let i = 1; i <= 200; i++) {
        let ans = "";

        let n = i;

        while (n > 0) {
            let rem = n % 26;
            if (rem == 0) {
                ans = "Z" + ans;
                n = Math.floor(n / 26) - 1;
            } else {
                ans = String.fromCharCode(rem - 1 + 65) + ans;
                n = Math.floor(n / 26);
            }
        }

        let column = $(`<div class="column-name colId-${i}" id="colCod-${ans}">${ans}</div>`);
        $(".column-name-container").append(column);
        let row = $(`<div class="row-name" id="rowId-${i}">${i}</div>`);
        $(".row-name-container").append(row);
    }

    for (let i = 1; i < 50; i++) {  //1048576 rows in ms excel
        let row = $(`<div class="cell-row"></div>`);
        for (let j = 1; j <= 50; j++) { //16384 col
            let colCode = $(`.colId-${j}`).attr("id").split("-")[1];
            let column = $(`<div class="input-cell" contenteditable="true" id = "row-${i}-col-${j}" data="code-${colCode}"></div>`);
            row.append(column);
        }
        $(".input-cell-container").append(row);
    }


    $(".align-icon").click(function () {
        $(".align-icon.selected").removeClass("selected");
        $(this).addClass("selected");
    });

    $(".style-icon").click(function () {
        $(this).toggleClass("selected");
    });

    $(".input-cell").click(function (e) {
        if (e.ctrlKey) {
            let [rowId, colId] = getRowCol(this);
            if (rowId > 1) {
                let topCellSelected = $(`#row-${rowId - 1}-col-${colId}`).hasClass("selected");
                if (topCellSelected) {
                    $(this).addClass("top-cell-selected");
                    $(`#row-${rowId - 1}-col-${colId}`).addClass("bottom-cell-selected");
                }
            }
            if (rowId < 100) {
                let bottomCellSelected = $(`#row-${rowId + 1}-col-${colId}`).hasClass("selected");
                if (bottomCellSelected) {
                    $(this).addClass("bottom-cell-selected");
                    $(`#row-${rowId + 1}-col-${colId}`).addClass("top-cell-selected");
                }
            }
            if (colId > 1) {
                let leftCellSelected = $(`#row-${rowId}-col-${colId - 1}`).hasClass("selected");
                if (leftCellSelected) {
                    $(this).addClass("left-cell-selected");
                    $(`#row-${rowId}-col-${colId - 1}`).addClass("right-cell-selected");
                }
            }
            if (colId < 100) {
                let rightCellSelected = $(`#row-${rowId}-col-${colId + 1}`).hasClass("selected");
                if (rightCellSelected) {
                    $(this).addClass("right-cell-selected");
                    $(`#row-${rowId}-col-${colId + 1}`).addClass("left-cell-selected");
                }
            }
        }
        else {
            $(".input-cell.selected").removeClass("selected");
        }
        $(this).addClass("selected");
        changeHeader(this);
    });

    function changeHeader(ele) {
        let [rowId, colId] = getRowCol(ele);
        let cellInfo = defaultProperties;
        if (cellData[selectedSheet][rowId] && cellData[selectedSheet][rowId][colId]) {
            cellInfo = cellData[selectedSheet][rowId][colId];
        }
        cellInfo["font-weight"] ? $(".icon-bold").addClass("selected") : $(".icon-bold").removeClass("selected");
        cellInfo["font-style"] ? $(".icon-italic").addClass("selected") : $(".icon-italic").removeClass("selected");
        cellInfo["text-decoration"] ? $(".icon-underline").addClass("selected") : $(".icon-underline").removeClass("selected");
        let alignment = cellInfo["text-align"];
        $(".align-icon.selected").removeClass("selected");
        $(".icon-align-" + alignment).addClass("selected");
        $(".background-color-picker").val(cellInfo["background-color"]);
        $(".text-color-picker").val(cellInfo["color"]);
        $(".font-family-selector").val(cellInfo["font-family"]);
        $(".font-family-selector").css("font-family", cellInfo["font-family"]);
        $(".font-size-selector").val(cellInfo["font-size"]);
    }

    const undoStack = [];
    const redoStack = [];

    // Wait for the DOM to load
    document.addEventListener("DOMContentLoaded", () => {
        // Attach input listeners to all editable cells
       document.querySelectorAll(".input-cell[contenteditable='true']").forEach(cell => {
            // Save the initial value
            cell.dataset.prev = cell.innerText;

            // Store value before change
            cell.addEventListener("focus", () => {
                cell.dataset.prev = cell.innerText;
            });

            // Track changes for undo
            cell.addEventListener("input", () => {
                undoStack.push({
                    row: cell.dataset.row,
                    col: cell.dataset.col,
                    value: cell.dataset.prev
                });

                redoStack.length = 0; // Clear redo stack on new edit
                cell.dataset.prev = cell.innerText;
            });
        });

        // Add click event listeners to undo/redo buttons
        document.getElementById('undo-btn').addEventListener('click', undo);
        document.getElementById('redo-btn').addEventListener('click', redo);
    });

    // Utility: get cell by row/col
    function getCell(row, col) {
        return document.querySelector(`[data-row='${row}'][data-col='${col}']`);
    }

    // Undo function
    function undo() {
        if (undoStack.length === 0) return;
        const last = undoStack.pop();
        const cell = getCell(last.row, last.col);
        redoStack.push({
            row: last.row,
            col: last.col,
            value: cell.innerText
        });
        cell.innerText = last.value;
        cell.dataset.prev = last.value;
        cell.focus();
    }

    // Redo function
    function redo() {
        if (redoStack.length === 0) return;
        const last = redoStack.pop();
        const cell = getCell(last.row, last.col);
        undoStack.push({
            row: last.row,
            col: last.col,
            value: cell.innerText
        });
        cell.innerText = last.value;
        cell.dataset.prev = last.value;
        cell.focus();
    }

    document.getElementById("file-menu").addEventListener("click", function (e) {
        e.stopPropagation();
        const dropdown = document.getElementById("file-dropdown");
        dropdown.classList.toggle("hidden");
    });

    // Hide dropdown if clicked outside
    document.addEventListener("click", function () {
        const dropdown = document.getElementById("file-dropdown");
        dropdown.classList.add("hidden");
    });
    $("#name-box").keypress(function (e) {
        if (e.which === 13) { // Enter key
            let input = $(this).val().toUpperCase().trim(); // e.g., "B2"
            let match = input.match(/^([A-Z]+)([0-9]+)$/);

            if (match) {
                let colLetters = match[1];
                let rowId = parseInt(match[2]);

                // Convert column letters to number
                let colId = 0;
                for (let i = 0; i < colLetters.length; i++) {
                    colId = colId * 26 + (colLetters.charCodeAt(i) - 64);
                }

                let $target = $(`#row-${rowId}-col-${colId}`);
                if ($target.length) {
                    $(".input-cell.selected").removeClass("selected");
                    $target.addClass("selected").attr("contenteditable", "true").focus();

                    // Scroll into view
                    let container = $(".input-cell-container");
                    container.scrollTop($target.position().top + container.scrollTop());
                    container.scrollLeft($target.position().left + container.scrollLeft());

                    // ✅ Update formula bar
                    let formula = $target.attr("data-formula");
                    if (formula) {
                        $(".formula-input").text(formula);
                    } else {
                        $(".formula-input").text($target.text());
                    }
                } else {
                    alert("Cell does not exist.");
                }
            } else {
                alert("Invalid cell reference (e.g., A1, B2).");
            }
        }
    });

    $(document).on("click", ".input-cell", function () {
        $(".input-cell.selected").removeClass("selected");
        $(this).addClass("selected");

        // Get cell row and column from id
        let idParts = $(this).attr("id").split("-");
        let rowId = parseInt(idParts[1]);
        let colId = parseInt(idParts[3]);

        // Convert column number to letters
        let colLetters = "";
        while (colId > 0) {
            let rem = (colId - 1) % 26;
            colLetters = String.fromCharCode(65 + rem) + colLetters;
            colId = Math.floor((colId - 1) / 26);
        }

        // Update Name Box and Formula Bar
        $("#name-box").val(colLetters + rowId);
        let formula = $(this).attr("data-formula");
        $(".formula-input").text(formula || $(this).text());


    });

    $(".input-cell").blur(function () {
    updateCell("text", $(this).text()); // Only keep this line
});

    $(".input-cell-container").scroll(function () {
        $(".column-name-container").scrollLeft(this.scrollLeft);
        $(".row-name-container").scrollTop(this.scrollTop);
    })

});

function getRowCol(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

function updateCell(property, value, defaultPossible) {
    $(".input-cell.selected").each(function () {
        $(this).css(property, value);
        let [rowId, colId] = getRowCol(this);
        if (cellData[selectedSheet][rowId]) {
            if (cellData[selectedSheet][rowId][colId]) {
                cellData[selectedSheet][rowId][colId][property] = value;
            } else {
                cellData[selectedSheet][rowId][colId] = { ...defaultProperties };
                cellData[selectedSheet][rowId][colId][property] = value;
            }
        } else {
            cellData[selectedSheet][rowId] = {};
            cellData[selectedSheet][rowId][colId] = { ...defaultProperties };
            cellData[selectedSheet][rowId][colId][property] = value;
        }
        if (defaultPossible && (JSON.stringify(cellData[selectedSheet][rowId][colId]) === JSON.stringify(defaultProperties))) {
            delete cellData[selectedSheet][rowId][colId];
            if (Object.keys(cellData[selectedSheet][rowId]).length == 0) {
                delete cellData[selectedSheet][rowId];
            }
        }
    });
    console.log(cellData);
}

$(".icon-bold").click(function () {
    if ($(this).hasClass("selected")) {
        updateCell("font-weight", "", true);
    } else {
        updateCell("font-weight", "bold", false);
    }
});

$(".icon-italic").click(function () {
    if ($(this).hasClass("selected")) {
        updateCell("font-style", "", true);
    } else {
        updateCell("font-style", "italic", false);
    }
});

$(".icon-underline").click(function () {
    if ($(this).hasClass("selected")) {
        updateCell("text-decoration", "", true);
    } else {
        updateCell("text-decoration", "underline", false);
    }
});

$(".icon-align-left").click(function () {
    if (!$(this).hasClass("selected")) {
        updateCell("text-align", "left", true);
    }
});

$(".icon-align-center").click(function () {
    if (!$(this).hasClass("selected")) {
        updateCell("text-align", "center", true);
    }
});

$(".icon-align-right").click(function () {
    if (!$(this).hasClass("selected")) {
        updateCell("text-align", "right", true);
    }
});

$(".color-fill-icon").click(function () {
    $(".background-color-picker").click();
});

$(".color-fill-text").click(function () {
    $(".text-color-picker").click();
});

$(".background-color-picker").change(function () {
    updateCell("background-color", $(this).val())
});

$(".text-color-picker").change(function () {
    updateCell("color", $(this).val())
});

$(".font-family-selector").change(function () {
    updateCell("font-family", $(this).val());
    $(".font-family-selector").css("font-family", $(this).val());
});

$(".font-size-selector").change(function () {
    updateCell("font-size", $(this).val());
});


function emptySheet() {
    let sheetInfo = cellData[selectedSheet];
    for (let i of Object.keys(sheetInfo)) {
        for (let j of Object.keys(sheetInfo[i])) {
            $(`#row-${i}-col-${j}`).text("");
            $(`#row-${i}-col-${j}`).css("background-color", "#ffffff");
            $(`#row-${i}-col-${j}`).css("color", "#000000");
            $(`#row-${i}-col-${j}`).css("text-align", "left");
            $(`#row-${i}-col-${j}`).css("font-weight", "");
            $(`#row-${i}-col-${j}`).css("font-style", "");
            $(`#row-${i}-col-${j}`).css("text-decoration", "");
            $(`#row-${i}-col-${j}`).css("font-family", "Noto Sans");
            $(`#row-${i}-col-${j}`).css("font-size", "14px");
        }
    }
}

function loadSheet() {
    let sheetInfo = cellData[selectedSheet];
    for (let i of Object.keys(sheetInfo)) {
        for (let j of Object.keys(sheetInfo[i])) {
            let cellInfo = cellData[selectedSheet][i][j];
            $(`#row-${i}-col-${j}`).text(cellInfo["text"]);
            $(`#row-${i}-col-${j}`).css("background-color", cellInfo["background-color"]);
            $(`#row-${i}-col-${j}`).css("color", cellInfo["color"]);
            $(`#row-${i}-col-${j}`).css("text-align", cellInfo["text-align"]);
            $(`#row-${i}-col-${j}`).css("font-weight", cellInfo["font-weight"]);
            $(`#row-${i}-col-${j}`).css("font-style", cellInfo["font-style"]);
            $(`#row-${i}-col-${j}`).css("text-decoration", cellInfo["text-decoration"]);
            $(`#row-${i}-col-${j}`).css("font-family", cellInfo["font-family"]);
            $(`#row-${i}-col-${j}`).css("font-size", cellInfo["font-size"]);
        }
    }
}

$(".icon-add").click(function () {
    emptySheet();
    $(".sheet-tab.selected").removeClass("selected");
    let sheetName = "Sheet" + (lastlyAddedSheet + 1);
    cellData[sheetName] = {};
    totalSheets += 1;
    lastlyAddedSheet += 1;
    selectedSheet = sheetName;
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">${sheetName}</div>`);
    addSheetEvents();
});


function addSheetEvents() {
    // Delegated contextmenu handler
    $(".sheet-tab-container").on("contextmenu", ".sheet-tab", function (e) {
        e.preventDefault();
        const clickedSheet = $(this); // ✅ Capture correct sheet reference
        selectSheet(clickedSheet.get(0));

        // Create context menu
        const menu = $(`<div class="sheet-options-modal">
                        <div class="sheet-rename">Rename</div>
                        <div class="sheet-delete">Delete</div>
                      </div>`).css({ top: e.pageY, left: e.pageX });

        // Delete handler
        menu.find(".sheet-delete").click(() => {
            if (Object.keys(cellData).length <= 1) {
                alert("Cannot delete last sheet");
                return;
            }

            // 1. Capture sheet name BEFORE deletion
            const sheetToDelete = selectedSheet;

            // 2. Select new sheet first
            const sheetTabs = $(".sheet-tab");
            const currIndex = sheetTabs.index(clickedSheet);
            const newSelected = currIndex > 0
                ? sheetTabs.eq(currIndex - 1)
                : sheetTabs.eq(currIndex + 1);

            if (newSelected.length) selectSheet(newSelected.get(0));

            // 3. Delete CORRECT data and DOM element
            delete cellData[sheetToDelete]; // ✅ Delete old data
            clickedSheet.remove(); // Remove correct DOM element
        });

        $("body").append(menu);
    });
}


function addSheetEvents() {
    $(".sheet-tab.selected").click(function () {
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });

    $(".sheet-tab.selected").contextmenu(function (e) {
        e.preventDefault();
        selectSheet(this);
        $(".sheet-options-modal").remove();

        // Get position of clicked sheet tab
        const tabRect = this.getBoundingClientRect();
        const menu = $(`<div class="sheet-options-modal">
                        <div class="sheet-rename">Rename</div>
                        <div class="sheet-delete">Delete</div>
                    </div>`);

        // Position menu above the sheet tab
        const menuTop = tabRect.top - 40; // 40px above the tab
        const menuLeft = tabRect.left + (tabRect.width / 2) - 60; // Centered

        menu.css({
            top: menuTop + "px",
            left: menuLeft + "px"
        });

        $("body").append(menu);

        // Rename functionality
        $(".sheet-rename").click(function () {
            $(".sheet-options-modal").remove();
            const renameModal = $(`<div class="sheet-rename-modal">
                                    <h4 class="modal-title">Rename Sheet To:</h4>
                                    <input type="text" class="new-sheet-name" placeholder="Sheet Name" />
                                    <div class="action-buttons">
                                        <div class="submit-button">Rename</div>
                                        <div class="cancel-button">Cancel</div>
                                    </div>
                                </div>`);

            // Position rename modal in center
            renameModal.css({
                top: "50%",
                left: "50%",
                transform: "translate(-50%, -50%)"
            });

            $("body").append(renameModal);

            $(".cancel-button").click(function () {
                $(".sheet-rename-modal").remove();
            });

            $(".submit-button").click(function () {
                const newSheetName = $(".new-sheet-name").val().trim();
                if (newSheetName) {
                    $(".sheet-tab.selected").text(newSheetName);
                    const prevName = selectedSheet;
                    cellData[newSheetName] = cellData[prevName];
                    delete cellData[prevName];
                    selectedSheet = newSheetName;
                }
                $(".sheet-rename-modal").remove();
            });
        });

        // Delete functionality
        $(".sheet-delete").click(function () {
            if (Object.keys(cellData).length > 1) {
                const currSheetName = selectedSheet;
                const currSheet = $(".sheet-tab.selected");
                const sheetTabs = $(".sheet-tab");
                const currIndex = sheetTabs.index(currSheet);

                // Select previous sheet if available, else next
                const newSelected = currIndex > 0 ?
                    sheetTabs.eq(currIndex - 1) :
                    sheetTabs.eq(currIndex + 1);

                newSelected.click();
                delete cellData[currSheetName];
                currSheet.remove();
            } else {
                alert("You must keep at least one sheet.");
            }
            $(".sheet-options-modal").remove();
        });

        // Close menu when clicking outside
        $(document).one("click", function (e) {
            if (!$(e.target).closest(".sheet-options-modal").length) {
                $(".sheet-options-modal").remove();
            }
        });
    });
}

// Handle container clicks
$(".container").click(function () {
    $(".sheet-options-modal").remove();
});
$(".icon-paste").click(function () {
    emptySheet();
    let [rowId, colId] = getRowCol($(".input-cell.selected")[0]);
    let rowDistance = rowId - selectedCells[0][0];
    let colDistance = colId - selectedCells[0][1];
    for (let cell of selectedCells) {
        let newRowId = cell[0] + rowDistance;
        let newColId = cell[1] + colDistance;
        if (!cellData[selectedSheet][newRowId]) {
            cellData[selectedSheet][newRowId] = {};
        }
        cellData[selectedSheet][newRowId][newColId] = { ...cellData[selectedSheet][cell[0]][cell[1]] };
        if (cut) {
            delete cellData[selectedSheet][cell[0]][cell[1]];
            if (Object.keys(cellData[selectedSheet][cell[0]]).length == 0) {
                delete cellData[selectedSheet][cell[0]];
            }
        }
    }
    if (cut) {
        cut = false;
        selectedCells = [];
    }
    loadSheet();
})

// Theme toggle logic
document.getElementById("theme-toggle").addEventListener("click", () => {
    const body = document.body;
    const currentTheme = body.getAttribute("data-theme") || "light";
    const newTheme = currentTheme === "light" ? "dark" : "light";
    body.setAttribute("data-theme", newTheme);
    document.getElementById("theme-toggle").textContent = newTheme === "light" ? "🌙" : "☀️";
});

// Set default theme
document.body.setAttribute("data-theme", "light");

