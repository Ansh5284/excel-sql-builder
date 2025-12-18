/* global document, Excel, Office, alasql, Sortable, console, setTimeout */

// STATE VARIABLES
let globalData = [];
let globalHeaders = []; // Real Excel Headers
let globalColTypes = {}; 
let virtualColumns = {}; // { "MyCol": "[Price]*[Qty]" }
let outputRangeAddress = null;
let activeFormulaInput = null; // Track focused formula input (in modal)

// FUNCTION DEFINITIONS
const FUNC_GROUPS = {
    GENERAL: ["COUNT", "DISTINCT"],
    MATH: ["SUM", "AVG", "MAX", "MIN", "ABS"],
    DATE: ["DAY", "MONTH", "YEAR"],
    STRING: ["LEN", "LOWER", "UPPER", "TRIM"]
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    if (typeof alasql === 'undefined') console.error("AlasQL not loaded!");
    else {
        alasql.options.mysql = true;
        alasql.fn.STRLEFT = (str, n) => typeof str === 'string' ? str.substring(0, n) : str;
        alasql.fn.STRRIGHT = (str, n) => typeof str === 'string' ? str.substring(str.length - n) : str;
    }

    initDragAndDrop();
    initLivePreviewListeners();

    document.getElementById("run-btn").onclick = runQuery;
    document.getElementById("btn-reset").onclick = resetUI;
    document.getElementById("btn-set-output").onclick = setOutputTarget;
    document.getElementById("date-helper").onchange = handleDatePick;
    
    // Modal Listeners
    document.getElementById("btn-cancel-col").onclick = closeModal;
    document.getElementById("btn-save-col").onclick = saveVirtualColumn;

    loadColumnsFromSelection();
  }
});

// ==========================================
// A. LOAD DATA
// ==========================================
async function loadColumnsFromSelection() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "address", "numberFormat"]); 
    await context.sync();

    if (!range.values || range.values.length < 2) {
        console.log("No data found in selection");
        return;
    }

    document.getElementById("current-table").innerText = "Range: " + range.address; 

    globalHeaders = range.values[0];
    const firstRowData = range.values[1];
    const firstRowFormat = range.numberFormat[1];

    globalHeaders.forEach((header, index) => {
        let fmt = firstRowFormat[index];
        if (typeof fmt === 'string' && fmt.toLowerCase().includes('d') && fmt.toLowerCase().includes('m')) {
            globalColTypes[header] = "DATE";
        } else if (typeof firstRowData[index] === 'number') {
            globalColTypes[header] = "NUMBER";
        } else {
            globalColTypes[header] = "TEXT";
        }
    });

    globalData = range.values.slice(1).map(row => {
      let obj = {};
      globalHeaders.forEach((h, i) => {
          let val = row[i];
          if (globalColTypes[h] === "DATE" && typeof val === 'number') {
              val = excelSerialToDate(val);
          }
          obj[h] = val;
      });
      return obj;
    });

    renderSourceList();
    updateSQLPreview();
  }).catch(handleError);
}

function excelSerialToDate(serial) {
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const year = date_info.getUTCFullYear();
    const month = String(date_info.getUTCMonth() + 1).padStart(2, '0');
    const day = String(date_info.getUTCDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// --- RENDER SOURCE LIST (SPLIT CHIPS) ---
function renderSourceList() {
  const list = document.getElementById("source-list");
  list.innerHTML = "";

  // 1. "Add Column" Chip
  let addBtn = document.createElement("div");
  addBtn.className = "item add-col-btn";
  addBtn.innerHTML = `<span class="item-name" style="padding:4px 12px;">+ Add Column</span>`;
  addBtn.onclick = openModal;
  list.appendChild(addBtn);

  // 2. Real Columns + Virtual Columns
  const allCols = [...globalHeaders, ...Object.keys(virtualColumns)];

  allCols.forEach(col => {
    const isVirtual = virtualColumns.hasOwnProperty(col);
    const type = globalColTypes[col] || "TEXT";
    
    // Create Chip Container
    let chip = document.createElement("div");
    chip.className = "item";
    chip.dataset.id = col;

    // A. Icon Side (Left)
    let icon = document.createElement("div");
    icon.className = "item-type";
    
    // Icon Logic
    if (isVirtual) {
        icon.innerText = "ƒx";
        icon.className += " type-calc";
        icon.title = "Calculated Column";
    } else {
        if (type === "NUMBER") { icon.innerText = "123"; icon.className += " type-num"; }
        else if (type === "DATE") { icon.innerText = "📅"; icon.className += " type-date"; }
        else { icon.innerText = "ABC"; icon.className += " type-txt"; }
        
        // Cycle Type on Click (Only for Real columns)
        icon.onclick = (e) => {
            e.stopPropagation();
            cycleColumnType(col);
        };
        icon.title = "Click to toggle type";
    }

    // B. Name Side (Right)
    let nameSpan = document.createElement("div");
    nameSpan.className = "item-name";
    nameSpan.innerText = col;

    // Click-to-Insert (into Modal Formula box)
    chip.onclick = function(e) {
        if (activeFormulaInput) {
            e.stopPropagation();
            insertTextAtCursor(activeFormulaInput, `[${col}]`);
        }
    };

    chip.appendChild(icon);
    chip.appendChild(nameSpan);
    list.appendChild(chip);
  });
}

function cycleColumnType(col) {
    const types = ["TEXT", "NUMBER", "DATE"];
    let current = globalColTypes[col] || "TEXT";
    let idx = types.indexOf(current);
    let next = types[(idx + 1) % types.length];
    
    updateColumnDataType(col, next);
    renderSourceList(); // Re-render to show new icon
    updateSQLPreview();
}

// --- MODAL LOGIC ---
function openModal() {
    document.getElementById("modal-add-col").style.display = "flex";
    const input = document.getElementById("new-col-formula");
    input.value = "";
    document.getElementById("new-col-name").value = "";
    input.focus();
    activeFormulaInput = input; // Set focus for click-to-insert
}

function closeModal() {
    document.getElementById("modal-add-col").style.display = "none";
    activeFormulaInput = null;
}

function saveVirtualColumn() {
    const name = document.getElementById("new-col-name").value.trim();
    const formula = document.getElementById("new-col-formula").value.trim();
    
    if (!name || !formula) {
        console.error("Name and Formula required");
        return;
    }
    
    // Check duplicates
    if (globalHeaders.includes(name) || virtualColumns[name]) {
        console.error("Column name exists");
        // Could add UI error message here
        return;
    }

    virtualColumns[name] = formula;
    globalColTypes[name] = "NUMBER"; // Default virtuals to Number (safest assumption for math)
    
    renderSourceList();
    closeModal();
}

function insertTextAtCursor(input, text) {
    const start = input.selectionStart;
    const end = input.selectionEnd;
    const val = input.value;
    input.value = val.substring(0, start) + text + val.substring(end);
    input.selectionStart = input.selectionEnd = start + text.length;
    input.focus();
    // Only update preview if NOT in modal (modal doesn't affect query yet)
    if (!document.getElementById("modal-add-col").style.display === "none") {
        updateSQLPreview();
    }
}

async function setOutputTarget() {
    await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        outputRangeAddress = range.address;
        document.getElementById("txt-output").value = outputRangeAddress.split("!")[1];
    }).catch(handleError);
}

// ==========================================
// B. DRAG & DROP & EVENTS
// ==========================================
function getDraggedColName(evt) {
    return evt.item.dataset.id || evt.item.innerText.trim();
}

function extractConfigFromEl(el) {
    const funcSel = el.querySelector(".func-select");
    const custInput = el.querySelector(".custom-func-input");
    const sortBtn = el.querySelector(".sort-btn");
    
    return {
        func: funcSel ? funcSel.value : "NONE",
        customVal: custInput ? custInput.value : "",
        sortDir: sortBtn ? sortBtn.dataset.dir : "ASC",
    };
}

function initDragAndDrop() {
    const source = document.getElementById("source-list");
    
    const sharedGroup = {
        name: 'shared',
        pull: function(to, from, dragEl, evt) {
            // Prevent dragging if Modal is open (prevent inserting into modal by drag)
            if (activeFormulaInput) return false;

            if (from.el.id === "source-list") return "clone";
            let isCtrl = false;
            if (evt.ctrlKey) isCtrl = true;
            if (evt.originalEvent && evt.originalEvent.ctrlKey) isCtrl = true;
            return isCtrl ? "clone" : true; 
        },
        put: true
    };

    const dropOps = {
        group: sharedGroup, 
        animation: 150,
        onSort: function(evt) { updateSQLPreview(); }
    };
    
    new Sortable(source, { 
        group: { name: 'shared', pull: 'clone', put: false }, 
        sort: false, 
        animation: 150 
    });

    const createHandler = (createFn) => {
        return function (evt) {
            createFn(getDraggedColName(evt), evt.item, extractConfigFromEl(evt.item)); 
            updateSQLPreview(); 
        };
    };

    new Sortable(document.getElementById("select-list"), { ...dropOps, onAdd: (evt) => createSelectRow(getDraggedColName(evt), evt.item, extractConfigFromEl(evt.item)) });
    new Sortable(document.getElementById("where-list"), { ...dropOps, onAdd: (evt) => createWhereRow(getDraggedColName(evt), evt.item, extractConfigFromEl(evt.item)) });
    new Sortable(document.getElementById("groupby-list"), { ...dropOps, onAdd: (evt) => createGroupRow(getDraggedColName(evt), evt.item, extractConfigFromEl(evt.item)) });
    new Sortable(document.getElementById("having-list"), { ...dropOps, onAdd: (evt) => createHavingRow(getDraggedColName(evt), evt.item, extractConfigFromEl(evt.item)) });
    new Sortable(document.getElementById("orderby-list"), { ...dropOps, onAdd: (evt) => createOrderRow(getDraggedColName(evt), evt.item, extractConfigFromEl(evt.item)) });
}

function initLivePreviewListeners() {
    document.querySelector('.container').addEventListener('change', function(e) {
        if (e.target.tagName === 'SELECT' || e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            updateSQLPreview();
        }
    });
    document.querySelector('.container').addEventListener('input', function(e) {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            updateSQLPreview();
        }
    });
}

function resetUI() {
    const restorePlaceholder = (id, text) => {
        document.getElementById(id).innerHTML = `<div class="placeholder">${text}</div>`;
    };
    restorePlaceholder("select-list", "Drop columns here...");
    restorePlaceholder("where-list", "Drop columns to filter...");
    restorePlaceholder("groupby-list", "Drop columns to group...");
    restorePlaceholder("having-list", "Drop columns to filter groups...");
    restorePlaceholder("orderby-list", "Drop columns to sort...");

    document.getElementById("chk-distinct").checked = false;
    document.getElementById("txt-limit").value = "";
    
    activeFormulaInput = null;
    virtualColumns = {}; // Clear virtuals
    renderSourceList();

    updateSQLPreview();
}

// ==========================================
// C. DATA TYPE HANDLING
// ==========================================
function updateColumnDataType(colName, newType) {
    console.log(`Converting column '${colName}' to ${newType}`);
    globalColTypes[colName] = newType;

    globalData.forEach(row => {
        let val = row[colName];
        if (val === undefined || val === null) return;

        if (newType === "NUMBER") {
            if (typeof val === 'string') {
                val = val.replace(/,/g, '').trim();
                if (val === '') val = 0;
            }
            let num = parseFloat(val);
            row[colName] = isNaN(num) ? 0 : num;
        } else if (newType === "TEXT") {
            row[colName] = String(val);
        } else if (newType === "DATE") {
            if (typeof val === 'number') {
                row[colName] = excelSerialToDate(val);
            } else {
                row[colName] = String(val);
            }
        }
    });
    // Note: No need to refresh Select Rows anymore as we removed Type Dropdowns there
}

function populateFuncDropdown(select, type) {
    const currentVal = select.value;
    select.innerHTML = ""; 
    let def = document.createElement("option");
    def.value = "NONE";
    def.innerText = "Raw";
    select.appendChild(def);

    function addGroup(label, funcs) {
        let group = document.createElement("optgroup");
        group.label = label;
        funcs.forEach(f => {
            let opt = document.createElement("option");
            opt.value = f;
            opt.innerText = f;
            group.appendChild(opt);
        });
        select.appendChild(group);
    }

    addGroup("General", FUNC_GROUPS.GENERAL);
    if (type === "NUMBER") addGroup("Math Functions", FUNC_GROUPS.MATH);
    else if (type === "DATE") addGroup("Date Functions", FUNC_GROUPS.DATE);
    else addGroup("String Functions", FUNC_GROUPS.STRING);

    let custom = document.createElement("option");
    custom.value = "CUSTOM";
    custom.innerText = "Other...";
    select.appendChild(custom);

    if(currentVal) select.value = currentVal;
    if(select.value === "") select.value = "NONE";
}

function applyConfig(row, config) {
    if (!config) return;
    const funcSel = row.querySelector(".func-select");
    const custInput = row.querySelector(".custom-func-input");
    
    if (funcSel && config.func) {
        funcSel.value = config.func;
        if (config.func === "CUSTOM" && config.customVal && custInput) {
            custInput.value = config.customVal;
            custInput.style.display = "inline-block";
        }
    }
}

// ==========================================
// D. ROW CREATION (CLEANUP)
// ==========================================

function createSelectRow(colName, targetElement, config) {
    let type = globalColTypes[colName] || "TEXT";
    const row = document.createElement("div");
    row.className = "select-row clause-select";
    row.dataset.id = colName;
    
    row.innerHTML = `<span class="col-label" title="${colName}">${colName}</span>`;
    
    // Func Dropdown (Standard)
    const select = document.createElement("select");
    select.className = "func-select";
    populateFuncDropdown(select, type);
    row.appendChild(select);

    const customInput = createCustomInput();
    row.appendChild(customInput);
    select.onchange = () => toggleCustomInput(select, customInput);

    addRemoveBtn(row);
    targetElement.replaceWith(row);
    applyConfig(row, config);
}

function createGroupRow(colName, targetElement, config) {
    const type = globalColTypes[colName] || "TEXT";
    const row = document.createElement("div");
    row.className = "select-row clause-group";
    row.dataset.id = colName;
    
    row.innerHTML = `<span class="col-label" title="${colName}">${colName}</span>`;

    const select = document.createElement("select");
    select.className = "func-select";
    populateFuncDropdown(select, type);
    row.appendChild(select);

    const customInput = createCustomInput();
    row.appendChild(customInput);
    select.onchange = () => toggleCustomInput(select, customInput);

    addRemoveBtn(row);
    targetElement.replaceWith(row);
    applyConfig(row, config);
}

function createWhereRow(colName, targetElement, config) {
    const type = globalColTypes[colName] || "TEXT";
    const row = document.createElement("div");
    row.className = "where-row clause-where";
    row.dataset.id = colName;
    
    const selectFunc = document.createElement("select");
    selectFunc.className = "func-select";
    selectFunc.style.width = "60px"; 
    populateFuncDropdown(selectFunc, type);
    row.appendChild(selectFunc);

    const customInput = createCustomInput(50, "Func");
    row.appendChild(customInput);
    selectFunc.onchange = () => toggleCustomInput(selectFunc, customInput);

    const nameDiv = document.createElement("div");
    nameDiv.className = "where-col-name";
    nameDiv.title = colName;
    nameDiv.innerText = colName;
    row.appendChild(nameDiv);

    addOperatorDropdown(row);
    addValueInput(row, type);
    addRemoveBtn(row); 
    targetElement.replaceWith(row);
    applyConfig(row, config);
}

function createHavingRow(colName, targetElement, config) {
    const type = globalColTypes[colName] || "TEXT";
    const row = document.createElement("div");
    row.className = "where-row clause-having";
    row.dataset.id = colName;
    
    const selectFunc = document.createElement("select");
    selectFunc.className = "func-select";
    selectFunc.style.width = "60px"; 
    populateFuncDropdown(selectFunc, type);
    row.appendChild(selectFunc);

    const customInput = createCustomInput(50, "Func");
    row.appendChild(customInput);
    selectFunc.onchange = () => toggleCustomInput(selectFunc, customInput);

    const nameDiv = document.createElement("div");
    nameDiv.className = "where-col-name";
    nameDiv.title = colName;
    nameDiv.innerText = colName;
    row.appendChild(nameDiv);

    addOperatorDropdown(row);
    addValueInput(row, type);
    addRemoveBtn(row); 
    targetElement.replaceWith(row);
    applyConfig(row, config);
}

function createOrderRow(colName, targetElement, config) {
    const type = globalColTypes[colName] || "TEXT";
    const row = document.createElement("div");
    row.className = "order-row clause-order";
    row.dataset.id = colName;
    
    const select = document.createElement("select");
    select.className = "func-select";
    select.style.width = "65px";
    populateFuncDropdown(select, type);
    row.appendChild(select);
    
    const customInput = createCustomInput(50);
    row.appendChild(customInput);
    select.onchange = () => toggleCustomInput(select, customInput);

    const span = document.createElement("span");
    span.className = "col-label";
    span.innerText = colName;
    span.style.flex = "1";
    row.appendChild(span);

    const btnSort = document.createElement("button");
    btnSort.className = "sort-btn asc";
    btnSort.innerText = "ASC";
    btnSort.dataset.dir = "ASC";
    btnSort.onclick = () => {
        if (btnSort.dataset.dir === "ASC") {
            btnSort.dataset.dir = "DESC";
            btnSort.innerText = "DESC";
            btnSort.className = "sort-btn desc";
        } else {
            btnSort.dataset.dir = "ASC";
            btnSort.innerText = "ASC";
            btnSort.className = "sort-btn asc";
        }
        updateSQLPreview();
    };
    row.appendChild(btnSort);

    addRemoveBtn(row);
    targetElement.replaceWith(row);

    applyConfig(row, config);
    if(config && config.sortDir) {
        btnSort.dataset.dir = config.sortDir;
        btnSort.innerText = config.sortDir;
        btnSort.className = `sort-btn ${config.sortDir.toLowerCase()}`;
    }
}

// --- SHARED UI HELPERS ---

function createCustomInput(width = 60, placeholder = "Func(@col)") {
    const input = document.createElement("input");
    input.type = "text";
    input.className = "custom-func-input";
    input.style.display = "none";
    input.style.width = width + "px";
    input.placeholder = placeholder;
    return input;
}

function toggleCustomInput(select, input) {
    input.style.display = (select.value === "CUSTOM") ? "inline-block" : "none";
}

function addOperatorDropdown(row) {
    const ops = ["=", ">", "<", ">=", "<=", "<>", "IN", "BETWEEN", "LIKE", "IS NULL", "IS NOT NULL"];
    let select = document.createElement("select");
    select.className = "where-op";
    ops.forEach(op => {
        let opt = document.createElement("option");
        opt.value = op;
        opt.innerText = op;
        select.appendChild(opt);
    });
    row.appendChild(select);
}

function addValueInput(row, type) {
    let valContainer = document.createElement("div");
    valContainer.className = "where-val-container";
    let input = document.createElement("input");
    input.type = "text";
    input.className = "where-val";
    valContainer.appendChild(input);

    if (type === "DATE") {
        let btn = document.createElement("span");
        btn.innerHTML = "📅";
        btn.className = "date-btn";
        btn.onclick = () => {
            activeDateInput = input;
            document.getElementById("date-helper").showPicker();
        };
        valContainer.appendChild(btn);
    }
    row.appendChild(valContainer);
}

function addRemoveBtn(row) {
    let remove = document.createElement("span");
    remove.className = "remove-btn";
    remove.innerText = "✖";
    remove.onclick = () => { row.remove(); updateSQLPreview(); };
    row.appendChild(remove);
}

let activeDateInput = null;
function handleDatePick(e) {
    if (!activeDateInput) return;
    if (activeDateInput.value.trim() !== "") activeDateInput.value += "," + e.target.value;
    else activeDateInput.value = e.target.value;
    updateSQLPreview();
}

// --- LOGIC HELPERS ---

function getColExpression(func, customVal, colName) {
    // 1. Check Virtual Columns substitution FIRST
    let baseCol = `[${colName}]`;
    if (virtualColumns[colName]) {
        baseCol = virtualColumns[colName]; // Wrap in parens for safety
    }

    if (func === "NONE") return baseCol;
    
    let expr = "";
    if (func === "CUSTOM") {
        let text = customVal.trim();
        if (!text) return baseCol; 
        
        // Replace placeholders
        if (text.includes("[@col]")) text = text.split("[@col]").join(baseCol);
        if (text.includes("@col")) {
            expr = text.split("@col").join(baseCol);
        } else {
            // If just text, wrap
            expr = `${text}(${baseCol})`;
        }
    } else {
        expr = `${func}(${baseCol})`;
    }
    expr = expr.replace(/\bLEFT\s*\(/gi, "STRLEFT(");
    expr = expr.replace(/\bRIGHT\s*\(/gi, "STRRIGHT(");
    return expr;
}

function buildConditionClause(row, aliasMap) {
    const funcSelect = row.querySelector(".func-select"); 
    const func = funcSelect ? funcSelect.value : "NONE";
    const customVal = row.querySelector(".custom-func-input").value;
    const colDiv = row.querySelector(".where-col-name");
    const col = colDiv ? colDiv.innerText : "";
    const opSelect = row.querySelector(".where-op");
    const op = opSelect ? opSelect.value : "=";
    const valInput = row.querySelector(".where-val");
    const valRaw = valInput ? valInput.value : "";
    
    let colStr = getColExpression(func, customVal, col);

    if (op === "IS NULL" || op === "IS NOT NULL") {
        return `${colStr} ${op}`;
    }
    else if (valRaw !== "") {
        let formattedVal = "";
        const cleanVal = valRaw.trim();
        const formatValue = (v) => {
            v = v.trim();
            if (v.startsWith("'") && v.endsWith("'")) return v; 
            if (!isNaN(v) && !isNaN(parseFloat(v))) return v;
            return `'${v}'`; 
        };
        if (op === "IN") {
            let parts = cleanVal.split(",").filter(v => v.trim() !== "").map(formatValue);
            if (parts.length > 0) formattedVal = `(${parts.join(",")})`;
        } else if (op === "BETWEEN") {
            let parts = cleanVal.split(",").filter(v => v.trim() !== "");
            if (parts.length >= 2) {
                formattedVal = `${formatValue(parts[0])} AND ${formatValue(parts[1])}`;
            }
        } else {
            formattedVal = formatValue(cleanVal);
        }
        if (formattedVal) return `${colStr} ${op} ${formattedVal}`;
    }
    return null;
}

// ==========================================
// D. QUERY GENERATION & EXECUTION
// ==========================================

function updateSQLPreview() {
    const sql = generateSQLQuery();
    document.getElementById("sql-preview").value = sql;
}

function generateSQLQuery() {
    let aliasMap = {}; 

    // 1. SELECT
    let colSqlParts = [];
    document.querySelectorAll(".clause-select").forEach(row => {
        const col = row.querySelector(".col-label").innerText;
        const func = row.querySelector(".func-select").value;
        const customVal = row.querySelector(".custom-func-input").value;
        
        const expr = getColExpression(func, customVal, col);
        let alias = col;
        if (func !== "NONE") {
            alias = (func === "CUSTOM") ? "Expr_" + col : func + "_" + col;
        }
        colSqlParts.push(`${expr} AS [${alias}]`);
        aliasMap[expr] = alias;
    });

    let cols = colSqlParts.join(", ");
    if (!cols) cols = "*";
    if (document.getElementById("chk-distinct").checked) cols = "DISTINCT " + cols;

    // 2. WHERE
    let whereClauses = [];
    document.querySelectorAll(".clause-where").forEach(row => {
        const clause = buildConditionClause(row, null);
        if (clause) whereClauses.push(clause);
    });
    const whereSql = whereClauses.length > 0 ? " WHERE " + whereClauses.join(" AND ") : "";

    // 3. GROUP BY
    let groupParts = [];
    document.querySelectorAll(".clause-group").forEach(row => {
        const col = row.querySelector(".col-label").innerText;
        const func = row.querySelector(".func-select").value;
        const customVal = row.querySelector(".custom-func-input").value;
        groupParts.push(getColExpression(func, customVal, col));
    });
    let groupSql = groupParts.length > 0 ? " GROUP BY " + groupParts.join(", ") : "";

    // 4. HAVING
    let havingClauses = [];
    document.querySelectorAll(".clause-having").forEach(row => {
        const clause = buildConditionClause(row, aliasMap);
        if (clause) havingClauses.push(clause);
    });
    const havingSql = havingClauses.length > 0 ? " HAVING " + havingClauses.join(" AND ") : "";

    // 5. ORDER BY
    let orderParts = [];
    document.querySelectorAll(".clause-order").forEach(row => {
        const col = row.querySelector(".col-label").innerText;
        const func = row.querySelector(".func-select").value;
        const customVal = row.querySelector(".custom-func-input").value;
        const direction = row.querySelector(".sort-btn").dataset.dir;

        const expr = getColExpression(func, customVal, col);

        if (aliasMap[expr]) {
             orderParts.push(`[${aliasMap[expr]}] ${direction}`);
        } else {
             orderParts.push(`${expr} ${direction}`);
        }
    });
    let orderSql = orderParts.length > 0 ? " ORDER BY " + orderParts.join(", ") : "";

    // 6. LIMIT
    const limitVal = document.getElementById("txt-limit").value;
    const limitSql = limitVal ? ` LIMIT ${limitVal}` : "";

    return `SELECT ${cols} FROM ? ${whereSql} ${groupSql} ${havingSql} ${orderSql} ${limitSql}`;
}

function runQuery() {
    let sql = generateSQLQuery();
    console.log("Executing SQL:", sql);

    try {
        const result = alasql(sql, [globalData]);
        console.log("Result Rows:", result.length);
        writeResult(result);
    } catch (err) {
        console.error("SQL Error:", err);
        document.getElementById("txt-output").value = "SQL Error: " + err.message;
    }
}

async function writeResult(data) {
    if (!data || data.length === 0) return;

    await Excel.run(async (context) => {
        let range;
        if (outputRangeAddress) {
            const sheetName = outputRangeAddress.split("!")[0];
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const startCell = outputRangeAddress.split("!")[1];
            let target = sheet.getRange(startCell);
            range = target.getResizedRange(data.length, Object.keys(data[0]).length - 1);
        } else {
            const sheet = context.workbook.worksheets.add();
            range = sheet.getRange("A1").getResizedRange(data.length, Object.keys(data[0]).length - 1);
            sheet.activate();
        }

        range.clear();
        
        const headers = Object.keys(data[0]);
        const values = data.map(row => headers.map(k => row[k]));
        range.values = [headers, ...values];
        range.format.autofitColumns();
        
        await context.sync();
    }).catch(handleError);
}

function handleError(error) {
    console.error("Excel Error: " + error);
}
