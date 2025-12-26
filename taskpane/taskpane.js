/* global document, Excel, Office, Sortable, console, setTimeout, initSqlJs */
// STATE VARIABLES
let globalData = [];
let globalHeaders = []; // Real Excel Headers
let globalColTypes = {}; 
let virtualColumns = {}; // { "MyCol": "[Price]*[Qty]" }
let outputRangeAddress = null;
let activeInput = null; // Track focused input GLOBALLY for modal insertion
let currentEditCol = null; 

// NEW: Registry to store all loaded tables for SQL execution
// Format: { "Table_Sheet1_A1": [ {Col1: Val, ...}, ... ], "Current Table": [...] }
let tableRegistry = {}; 

// FUNCTION DEFINITIONS
const FUNC_GROUPS = {
    GENERAL: ["COUNT", "DISTINCT"],
    MATH: ["SUM", "AVG", "MAX", "MIN", "ABS"],
    DATE: ["DAY", "MONTH", "YEAR"],
    STRING: ["LEN", "LOWER", "UPPER", "TRIM"]
};

/**
 * Executes a SQL query using SQLite (sql.js)
 * @param {string} query - The SQL query string
 * @param {Object} tableData - An object mapping table names to data arrays.
 */
async function executeSqlite(query, tableData) {
    // 1. Initialize SQLite
    // Ensure the WASM file is accessible. Ideally, place sql-wasm.wasm in your project
    // or use the CDN link below.
    const config = {
        locateFile: filename => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${filename}`
    };
    
    const SQL = await initSqlJs(config);
    const db = new SQL.Database();

    // REGISTER CUSTOM FUNCTIONS to match your existing logic
    db.create_function("STRLEFT", (str, n) => typeof str === 'string' ? str.substring(0, n) : str);
    db.create_function("STRRIGHT", (str, n) => typeof str === 'string' ? str.substring(str.length - n) : str);

    try {
        // 2. Load Data into Tables
        for (const [tableName, rows] of Object.entries(tableData)) {
            if (!rows || rows.length === 0) continue;

            const columns = Object.keys(rows[0]);
            
            // Create Table
            // SQLite is flexible with types, so we don't strictly need to define them here for this use case
            const createSql = `CREATE TABLE [${tableName}] (${columns.map(c => `[${c}]`).join(', ')});`;
            db.run(createSql);

            // Insert Data (Prepare once, run many)
            const placeholders = columns.map(() => '?').join(',');
            const insertSql = `INSERT INTO [${tableName}] VALUES (${placeholders})`;
            
            const stmt = db.prepare(insertSql);
            try {
                rows.forEach(row => {
                    const values = columns.map(col => row[col]);
                    stmt.run(values);
                });
            } finally {
                stmt.free();
            }
        }

        // 3. Execute the actual User Query
        const resultSets = db.exec(query);
        
        if (!resultSets || resultSets.length === 0) return [];

        // 4. Convert SQLite result format back to JSON (Array of Objects)
        const columns = resultSets[0].columns;
        const values = resultSets[0].values;

        const jsonResult = values.map(row => {
            let obj = {};
            columns.forEach((col, index) => {
                obj[col] = row[index];
            });
            return obj;
        });

        return jsonResult;

    } catch (error) {
        console.error("SQLite Error:", error);
        throw error;
    } finally {
        // 5. Close DB to free memory
        db.close();
    }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // REMOVED: AlasQL checks and custom function registration
    // (These are now handled inside executeSqlite)

    initDragAndDrop();
    initLivePreviewListeners();

    document.getElementById("run-btn").onclick = runQuery;
    document.getElementById("btn-reset").onclick = resetUI;
    document.getElementById("btn-set-output").onclick = setOutputTarget;
    document.getElementById("date-helper").onchange = handleDatePick;
    
    // Stage 3 Listener
    document.getElementById("btn-open-advanced").onclick = openAdvancedEditor;

    // Modal Listeners
    document.getElementById("btn-cancel-col").onclick = closeModal;
    document.getElementById("btn-close-x").onclick = closeModal;
    document.getElementById("btn-save-col").onclick = saveVirtualColumn;
    document.getElementById("btn-add-rule").onclick = addConditionRow;

    window.switchModalTab = switchModalTab; // Global expose for HTML

    loadColumnsFromSelection();
  }
});

// --- STAGE 3: ADVANCED EDITOR ---
let dialog = null;

function openAdvancedEditor() {
    const fullUrl = "https://ansh5284.github.io/excel-sql-builder/dialog/dialog.html";

    Office.context.ui.displayDialogAsync(fullUrl, { height: 80, width: 80, displayInIframe: true }, 
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processDialogMessage);
                
                // Send current schema to dialog
                setTimeout(() => {
                    const message = JSON.stringify({
                        type: "init",
                        schema: globalHeaders,
                        virtuals: virtualColumns,
                        colTypes: globalColTypes
                    });
                    dialog.messageChild(message);
                }, 1000); 
                setTimeout(() => {
                    // Retry just in case
                    if(dialog) {
                        const message = JSON.stringify({
                            type: "init",
                            schema: globalHeaders,
                            virtuals: virtualColumns,
                            colTypes: globalColTypes
                        });
                        dialog.messageChild(message);
                    }
                }, 3000);
            }
        }
    );
}

function processDialogMessage(arg) {
    const message = JSON.parse(arg.message);
    
    if (message.action === "close") {
        dialog.close();
        dialog = null;
    }
    else if (message.action === "addTable") {
        fetchTableForDialog();
    }
    else if (message.action === "runQuery") {
        console.log (message.sql, message.tables)
        executeAdvancedQuery(message.sql,  message.tables);
    }
}

// FIX: Execution Logic for Multi-Table (Migrated to SQLite)
async function executeAdvancedQuery(sql, tables) {
    console.log("Executing Advanced Query:", sql);
    
    // 1. Determine which tables to use
    const tablesToUse = (tables && Array.isArray(tables) && tables.length > 0) 
        ? tables 
        : Object.keys(tableRegistry);

    // 2. Build Data Mapping for SQLite
    let dataMapping = {};
    tablesToUse.forEach(tableName => {
        if (tableRegistry[tableName]) {
            dataMapping[tableName] = tableRegistry[tableName];
        } else {
             console.warn(`Warning: Data for table [${tableName}] is missing from registry.`);
        }
    });

    // 3. Execute
    try {
        let result = await executeSqlite(sql, dataMapping);
        console.log("Result Rows:", result.length);
        writeResult(result);
    } catch (err) {
        console.error("SQL Error:", err);
        document.getElementById("txt-output").value = "Error: " + err.message;
    }
}

// FIX: Helper to fetch selection and STORE IT locally
async function fetchTableForDialog() {
    await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address", "numberFormat"]);
        await context.sync();

        if (!range.values || range.values.length < 2) return;

        // Generate Name
        let rawName = range.address.split("!").pop().replace(/[^a-zA-Z0-9]/g, '');
        let tableName = "Table_" + rawName; 

        // 1. Process Headers
        const rawHeaders = range.values[0];
        let headers = [];
        let headerCounts = {};
        rawHeaders.forEach((h, i) => {
            let clean = (h === null || h === undefined || h === "") ? `Column_${i+1}` : String(h);
            if (headerCounts[clean]) {
                let count = headerCounts[clean]++;
                clean = `${clean}_${count}`;
            } else {
                headerCounts[clean] = 1;
            }
            headers.push(clean);
        });

        // 2. Process Data & Types
        const firstRowData = range.values[1];
        const firstRowFormat = range.numberFormat[1];
        let colTypes = {};
        
        // Detect types
        headers.forEach((header, index) => {
            let fmt = firstRowFormat[index];
            if (typeof fmt === 'string' && fmt.toLowerCase().includes('d') && fmt.toLowerCase().includes('m')) {
                colTypes[header] = "DATE";
            } else if (typeof firstRowData[index] === 'number') {
                colTypes[header] = "NUMBER";
            } else {
                colTypes[header] = "TEXT";
            }
        });

        // Build Data Array
        let tableData = range.values.slice(1).map(row => {
            let obj = {};
            headers.forEach((h, i) => {
                let val = row[i];
                let type = colTypes[h];

                if (type === "NUMBER") {
                    val = cleanNumber(val); 
                } else if (type === "DATE") {
                    if (typeof val === 'number') val = excelSerialToDate(val);
                    else val = String(val || "");
                } else {
                    val = (val === null || val === undefined) ? "" : String(val);
                }
                obj[h] = val;
            });
            return obj;
        });

        // 3. STORE IN REGISTRY
        tableRegistry[tableName] = tableData;

        // 4. Send Schema to Dialog
        if (dialog) {
            const msg = JSON.stringify({
                type: "addTableSuccess",
                tableName: tableName,
                schema: headers,
                colTypes: colTypes
            });
            dialog.messageChild(msg);
        }
    }).catch(console.error);
}

// --- AGGRESSIVE NUMBER CLEANER ---
function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (val === null || val === undefined || val === '') return 0;
    let str = String(val).replace(/[^0-9.-]/g, '');
    let num = parseFloat(str);
    return isNaN(num) ? 0 : num;
}

// ==========================================
// A. LOAD DATA (Task Pane Main Load)
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

    // 1. SANITIZE HEADERS
    const rawHeaders = range.values[0];
    globalHeaders = [];
    let headerCounts = {};

    rawHeaders.forEach((h, i) => {
        let cleanHeader = (h === null || h === undefined || h === "") ? `Column_${i+1}` : String(h);
        if (headerCounts[cleanHeader]) {
            let count = headerCounts[cleanHeader]++;
            cleanHeader = `${cleanHeader}_${count}`;
        } else {
            headerCounts[cleanHeader] = 1;
        }
        globalHeaders.push(cleanHeader);
    });

    const firstRowData = range.values[1];
    const firstRowFormat = range.numberFormat[1];

    // 2. DETECT TYPES
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

    // 3. BUILD DATA
    globalData = range.values.slice(1).map(row => {
      let obj = {};
      globalHeaders.forEach((h, i) => {
          let val = row[i];
          const type = globalColTypes[h];

          if (type === "NUMBER") {
              val = cleanNumber(val); 
          } else if (type === "DATE") {
              if (typeof val === 'number') {
                  val = excelSerialToDate(val);
              } else {
                  val = String(val || "");
              }
          } else {
              val = (val === null || val === undefined) ? "" : String(val);
          }
          
          obj[h] = val;
      });
      return obj;
    });

    // STORE IN REGISTRY (The fix for "Current Table")
    tableRegistry["Current Table"] = globalData;

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

// --- RENDER SOURCE LIST ---
function renderSourceList() {
  const list = document.getElementById("source-list");
  list.innerHTML = "";

  let addBtn = document.createElement("div");
  addBtn.className = "item add-col-btn";
  addBtn.innerHTML = `<span class="item-name" style="padding:4px 12px;">+ Add Column</span>`;
  addBtn.onclick = () => openModal(null);
  list.appendChild(addBtn);

  const allCols = [...globalHeaders, ...Object.keys(virtualColumns)];

  allCols.forEach(col => {
    const isVirtual = virtualColumns.hasOwnProperty(col);
    const type = globalColTypes[col] || "TEXT";
    
    let chip = document.createElement("div");
    chip.className = "item";
    chip.dataset.id = col;

    let icon = document.createElement("div");
    icon.className = "item-type";
    
    if (isVirtual) {
        icon.innerText = "ƒx";
        icon.className += " type-calc";
        icon.title = "Calculated Column - Click to Edit";
        // CLICK TO EDIT LOGIC
        icon.onclick = (e) => {
            e.stopPropagation();
            openModal(col); // Edit this column
        };
    } else {
        if (type === "NUMBER") { icon.innerText = "123"; icon.className += " type-num"; }
        else if (type === "DATE") { icon.innerText = "📅"; icon.className += " type-date"; }
        else { icon.innerText = "ABC"; icon.className += " type-txt"; }
        
        icon.onclick = (e) => {
            e.stopPropagation();
            cycleColumnType(col);
        };
        icon.title = "Click to toggle type";
    }

    let nameSpan = document.createElement("div");
    nameSpan.className = "item-name";
    nameSpan.innerText = col;

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
    renderSourceList(); 
    updateSQLPreview();
}

// --- MODAL LOGIC (TABS & CONDITIONAL) ---

function switchModalTab(tabName) {
    // Reset classes
    document.getElementById("tab-btn-formula").classList.remove("active");
    document.getElementById("tab-btn-cond").classList.remove("active");
    document.getElementById("view-formula").style.display = "none";
    document.getElementById("view-cond").style.display = "none";

    // Activate
    document.getElementById("tab-btn-" + tabName).classList.add("active");
    document.getElementById("view-" + tabName).style.display = "block";
    
    // Set active input for click-to-insert
    if (tabName === 'formula') {
        const formulaInput = document.getElementById("new-col-formula");
        formulaInput.focus();
        activeInput = formulaInput;
    } else {
        activeInput = null;
    }
}

function openModal(colName) {
    document.getElementById("modal-add-col").style.display = "flex";
    const nameInput = document.getElementById("new-col-name");
    const formulaInput = document.getElementById("new-col-formula");
    const saveBtn = document.getElementById("btn-save-col");

    // 1. POPULATE CHIP LIST IN MODAL (With Icons)
    const modalChipList = document.getElementById("modal-chip-list");
    modalChipList.innerHTML = "";
    
    const allCols = [...globalHeaders, ...Object.keys(virtualColumns)];
    allCols.forEach(col => {
        // Don't show self if editing
        if (col === colName) return;

        let chip = document.createElement("div");
        chip.className = "item";
        chip.style.margin = "2px";
        
        // Add Icon for consistency and better visuals
        const type = globalColTypes[col] || "TEXT";
        const isVirtual = virtualColumns.hasOwnProperty(col);
        let iconHtml = "";
        if (isVirtual) {
            iconHtml = `<div class="item-type type-calc" style="min-width:25px; font-size:10px;">ƒx</div>`;
        } else if (type === "NUMBER") { 
            iconHtml = `<div class="item-type type-num" style="min-width:25px; font-size:10px;">123</div>`;
        } else if (type === "DATE") { 
            iconHtml = `<div class="item-type type-date" style="min-width:25px; font-size:10px;">📅</div>`;
        } else { 
            iconHtml = `<div class="item-type type-txt" style="min-width:25px; font-size:10px;">ABC</div>`;
        }

        chip.innerHTML = `${iconHtml}<span class="item-name" style="padding:2px 8px;">${col}</span>`;
        
        // CLICK TO INSERT INTO ACTIVE INPUT
        chip.onclick = (e) => {
            e.stopPropagation();
            if (activeInput) {
                insertTextAtCursor(activeInput, `[${col}]`);
            } else {
                // Fallback: If no input focused, default to formula box if it's visible
                if(document.getElementById("view-formula").style.display !== 'none'){
                    insertTextAtCursor(formulaInput, `[${col}]`);
                }
            }
        };
        modalChipList.appendChild(chip);
    });


    // 2. SETUP FIELDS
    document.getElementById("cond-rows-container").innerHTML = "";
    document.getElementById("cond-else").value = "";
    document.getElementById("cond-else").onfocus = (e) => activeInput = e.target;

    if (colName) {
        // EDIT MODE
        currentEditCol = colName;
        nameInput.value = colName;
        nameInput.disabled = true; 
        
        // Check if formula looks like a CASE statement
        if (virtualColumns[colName].trim().toUpperCase().startsWith("CASE")) {
             formulaInput.value = virtualColumns[colName];
             switchModalTab('formula');
        } else {
             formulaInput.value = virtualColumns[colName];
             switchModalTab('formula');
        }
        
        saveBtn.innerText = "Update Column";
    } else {
        // NEW MODE
        currentEditCol = null;
        nameInput.value = "";
        nameInput.disabled = false;
        formulaInput.value = "";
        saveBtn.innerText = "Add Column";
        switchModalTab('formula');
        addConditionRow(); // Start with one row
    }
    
    // Focus tracking for formula box
    formulaInput.onfocus = (e) => activeInput = e.target;
    formulaInput.focus();
    activeInput = formulaInput;
}

function closeModal() {
    document.getElementById("modal-add-col").style.display = "none";
    activeInput = null;
    currentEditCol = null;
}

// --- CONDITIONAL BUILDER LOGIC ---
function addConditionRow() {
    const container = document.getElementById("cond-rows-container");
    const div = document.createElement("div");
    div.className = "cond-row";
    
    // Using Text Inputs for Columns now to allow multiple columns / math
    div.innerHTML = `
        <span class="cond-if">IF</span>
        <input type="text" class="cond-input-col" placeholder="Column/Expr">
        <select class="cond-op" title="Operator">
            <option value="=">=</option>
            <option value=">">></option>
            <option value="<"><</option>
            <option value=">=">>=</option>
            <option value="<="><=</option>
            <option value="!=">!=</option>
            <option value="LIKE">LIKE</option>
        </select>
        <input type="text" class="cond-val" placeholder="Value">
        <span class="cond-then">THEN</span>
        <input type="text" class="cond-res" placeholder="Result">
        <span class="cond-del" title="Remove Rule">✖</span>
    `;
    
    // Add focus listeners for chip insertion
    const inputs = div.querySelectorAll("input");
    inputs.forEach(inp => {
        inp.onfocus = (e) => activeInput = e.target;
    });

    // Add delete listener
    div.querySelector(".cond-del").onclick = () => div.remove();

    container.appendChild(div);
    
    // Auto-focus the first input (Column/Expr) of the new row so user can immediately click a chip
    const firstInput = div.querySelector(".cond-input-col");
    if(firstInput) {
        firstInput.focus();
        activeInput = firstInput;
    }
}

function buildCaseStatement() {
    const rows = document.querySelectorAll(".cond-row");
    if (rows.length === 0) return "";

    let sql = "CASE ";
    rows.forEach(row => {
        let col = row.querySelector(".cond-input-col").value.trim();
        const op = row.querySelector(".cond-op").value;
        let val = row.querySelector(".cond-val").value.trim();
        let res = row.querySelector(".cond-res").value.trim();

        if(!col) return; // Skip empty rows

        // Auto-quote strings if they aren't numbers or columns
        if (isNaN(val) && !val.startsWith("'") && !val.startsWith("[")) val = `'${val}'`;
        if (isNaN(res) && !res.startsWith("'") && !res.startsWith("[")) res = `'${res}'`;

        sql += `WHEN ${col} ${op} ${val} THEN ${res} `;
    });

    let elseVal = document.getElementById("cond-else").value.trim();
    if (!elseVal) elseVal = "NULL";
    else if (isNaN(elseVal) && !elseVal.startsWith("'") && !elseVal.startsWith("[")) elseVal = `'${elseVal}'`;

    sql += `ELSE ${elseVal} END`;
    return sql;
}

function saveVirtualColumn() {
    const name = document.getElementById("new-col-name").value.trim();
    let formula = "";

    if (document.getElementById("tab-btn-cond").classList.contains("active")) {
        formula = buildCaseStatement();
    } else {
        formula = document.getElementById("new-col-formula").value.trim();
    }
    
    if (!name || !formula) {
        console.error("Name and Formula required");
        return;
    }
    
    if (!currentEditCol && (globalHeaders.includes(name) || virtualColumns[name])) {
        console.error("Column name exists");
        return;
    }

    virtualColumns[name] = formula;
    globalColTypes[name] = "NUMBER"; // Default
    
    renderSourceList();
    closeModal();
    updateSQLPreview();
}

function insertTextAtCursor(input, text) {
    if (!input) return;
    const start = input.selectionStart;
    const end = input.selectionEnd;
    const val = input.value;
    input.value = val.substring(0, start) + text + val.substring(end);
    input.selectionStart = input.selectionEnd = start + text.length;
    input.focus();
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
            if (activeInput) return false; // Prevent drag if editing modal
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
    virtualColumns = {}; 
    renderSourceList();

    updateSQLPreview();
}

// ==========================================
// C. DATA TYPE HANDLING (Strict)
// ==========================================
function updateColumnDataType(colName, newType) {
    console.log(`Converting column '${colName}' to ${newType}`);
    globalColTypes[colName] = newType;

    globalData.forEach(row => {
        let val = row[colName];
        if (newType === "NUMBER") {
            val = cleanNumber(val); 
            row[colName] = val;
        } else if (newType === "TEXT") {
            row[colName] = (val === null || val === undefined) ? "" : String(val);
        } else if (newType === "DATE") {
            if (typeof val === 'number') {
                row[colName] = excelSerialToDate(val);
            } else {
                row[colName] = String(val || "");
            }
        }
    });
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
// D. ROW CREATION
// ==========================================

function createSelectRow(colName, targetElement, config) {
    let type = globalColTypes[colName] || "TEXT";
    const row = document.createElement("div");
    row.className = "select-row clause-select";
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
    let baseCol = `[${colName}]`;
    if (virtualColumns[colName]) {
        baseCol = virtualColumns[colName]; // NO PARENTHESES WRAPPER
        
        // FIX: Recursively replace other Virtual Columns inside this formula
        Object.keys(virtualColumns).forEach(key => {
            // Check if this formula uses another virtual column (e.g., [Total GST])
            if (key !== colName && baseCol.includes(`[${key}]`)) {
                // Replace [Total GST] with (Its Formula) to ensure SQL sees raw logic
                baseCol = baseCol.split(`[${key}]`).join(`(${virtualColumns[key]})`);
            }
        });
    }

    if (func === "NONE") return baseCol;
    
    let expr = "";
    if (func === "CUSTOM") {
        let text = customVal.trim();
        if (!text) return baseCol; 
        
        if (text.includes("[@col]")) text = text.split("[@col]").join(baseCol);
        if (text.includes("@col")) {
            expr = text.split("@col").join(baseCol);
        } else {
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

    // MIGRATION: Changed FROM ? to FROM [Current Table] for SQLite
    return `SELECT ${cols} FROM [Current Table] ${whereSql} ${groupSql} ${havingSql} ${orderSql} ${limitSql}`;
}

// FIX: Execution Logic for Single Table (Migrated to SQLite)
async function runQuery() {
    let sql = generateSQLQuery();
    console.log("Executing SQL:", sql);

    try {
        // Build Data Mapping for SQLite
        // Key "Current Table" must match the name used in generateSQLQuery
        const dataMapping = { "Current Table": globalData };
        
        let result = await executeSqlite(sql, dataMapping);
        console.log("Result Rows:", result.length);
        writeResult(result);
    } catch (err) {
        console.error("SQL Error:", err);
        document.getElementById("txt-output").value = "SQL Error: " + err.message;
    }
}

async function writeResult(data) {
    if (!data || data.length === 0) return;

    // Apply Rounding Preference to Output Data
    // Loop through data and if value is a float, round to 2 digits
    data.forEach(row => {
        Object.keys(row).forEach(key => {
            let val = row[key];
            if (typeof val === 'number' && !Number.isInteger(val)) {
                row[key] = parseFloat(val.toFixed(2));
            }
        });
    });

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
