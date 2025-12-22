/* global Office, document, console */

// Local state
let currentData = {
    tables: {}, 
    virtuals: {},
    joins: [] 
};

let tablesOnCanvas = []; 
let activeInput = null; 
let currentEditCol = null;
let currentJoinId = null;

// Drag State for Moving Cards
let dragSrcCard = null;
let dragOffsetX = 0;
let dragOffsetY = 0;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('btn-close').onclick = () => Office.context.ui.messageParent(JSON.stringify({ action: "close" }));
        
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent);

        // UI Event Listeners
        document.getElementById('btn-add-table').onclick = () => Office.context.ui.messageParent(JSON.stringify({ action: "addTable" }));
        document.getElementById('btn-add-calc').onclick = () => openModal(null);
        document.getElementById('btn-cancel-col').onclick = closeModal;
        document.getElementById('btn-close-x').onclick = closeModal;
        document.getElementById('btn-save-col').onclick = saveVirtualColumn;
        document.getElementById('btn-add-rule').onclick = addConditionRow;

        document.getElementById('btn-delete-join').onclick = deleteCurrentJoin;
        document.getElementById('btn-save-join').onclick = saveCurrentJoin;
        
        // Apply Button (Run)
        document.getElementById('btn-apply').onclick = () => {
            const sql = document.getElementById('sql-editor-area').value;
            Office.context.ui.messageParent(JSON.stringify({ action: "runQuery", sql: sql }));
        };

        // Canvas Drop
        const canvasZone = document.getElementById('visual-drop-zone');
        canvasZone.addEventListener('dragover', (e) => { e.preventDefault(); });
        canvasZone.addEventListener('drop', handleCanvasDrop);

        // SQL Editor Drop
        const sqlArea = document.getElementById('sql-editor-area');
        sqlArea.addEventListener('dragover', (e) => { e.preventDefault(); sqlArea.style.border="2px dashed #0078d4"; });
        sqlArea.addEventListener('dragleave', (e) => { sqlArea.style.border="none"; });
        sqlArea.addEventListener('drop', handleSqlDrop);

        window.switchTab = switchTab;
        window.switchModalTab = switchModalTab;
        window.onresize = renderJoins;
    }
});

function onMessageFromParent(arg) {
    const message = JSON.parse(arg.message);
    if (message.type === "init") {
        currentData.tables["Current Table"] = { schema: message.schema, colTypes: message.colTypes || {} };
        currentData.virtuals = message.virtuals || {};
        renderIngredients();
    }
    if (message.type === "addTableSuccess") {
        currentData.tables[message.tableName] = { schema: message.schema, colTypes: message.colTypes };
        renderIngredients();
    }
}

// --- RENDER SIDEBAR ---
function renderIngredients() {
    const container = document.getElementById('ingredients-list');
    container.innerHTML = "";
    
    Object.keys(currentData.tables).forEach(tableName => {
        const group = document.createElement('div');
        group.className = 'table-group';
        
        const header = document.createElement('div');
        header.className = 'table-header';
        header.innerHTML = `<span>▼ ${tableName}</span>`;
        header.draggable = true;
        header.ondragstart = (e) => { e.dataTransfer.setData("application/json", JSON.stringify({ type: "table", name: tableName })); };
        group.appendChild(header);
        
        currentData.tables[tableName].schema.forEach(col => {
            const type = currentData.tables[tableName].colTypes[col] || "TEXT";
            group.appendChild(createChip(col, type, false, tableName));
        });
        container.appendChild(group);
    });
    
    const virtualKeys = Object.keys(currentData.virtuals);
    if (virtualKeys.length > 0) {
        const vGroup = document.createElement('div');
        vGroup.className = 'table-group';
        vGroup.innerHTML = `<div class="table-header" style="color:#813a7c">▼ Calculated</div>`;
        virtualKeys.forEach(key => {
            const chip = createChip(key, "CALC", true);
            vGroup.appendChild(chip);
        });
        container.appendChild(vGroup);
    }
}

function createChip(colName, type, isVirtual, tableName="") {
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.draggable = true;
    chip.ondragstart = (e) => e.dataTransfer.setData("text/plain", tableName ? `[${tableName}].[${colName}]` : `[${colName}]`);
    
    let iconClass = "type-txt", iconText = "ABC";
    if (type==="NUMBER") { iconClass="type-num"; iconText="123"; }
    else if (type==="DATE") { iconClass="type-date"; iconText="📅"; }
    else if (isVirtual || type === "CALC") { iconClass="type-calc"; iconText="ƒx"; }
    
    chip.innerHTML = `<div class="chip-icon ${iconClass}">${iconText}</div><div class="chip-name">${colName}</div>`;
    
    // FIX: Type Toggle / Edit Formula Logic
    const iconDiv = chip.querySelector('.chip-icon');
    if (isVirtual || type === "CALC") {
        iconDiv.onclick = (e) => { e.stopPropagation(); openModal(colName); };
        iconDiv.title = "Edit Formula";
    } else {
        // FIX: Ensure this updates the correct table's type definition
        iconDiv.onclick = (e) => { e.stopPropagation(); toggleType(colName, tableName); };
        iconDiv.title = "Toggle Type";
    }
    return chip;
}

// FIX: Toggle Type handles specific table
function toggleType(colName, tableName) {
    const types = ["TEXT", "NUMBER", "DATE"];
    // Fallback to "Current Table" if tableName missing (legacy)
    const targetTable = tableName || "Current Table";
    
    if (currentData.tables[targetTable]) {
        let current = currentData.tables[targetTable].colTypes[colName] || "TEXT";
        let idx = types.indexOf(current);
        let next = types[(idx + 1) % types.length];
        currentData.tables[targetTable].colTypes[colName] = next;
        renderIngredients();
    }
}

// --- VISUAL CANVAS LOGIC ---
function handleCanvasDrop(e) {
    e.preventDefault();
    const rawData = e.dataTransfer.getData("application/json");
    if (!rawData) return;
    const data = JSON.parse(rawData);
    
    if (data.type === "table") {
        addTableToCanvas(data.name, e.clientX, e.clientY);
    } else if (data.type === "move_card") {
        // FIX: Move Existing Card
        const card = tablesOnCanvas.find(t => t.id === data.id);
        if (card) {
            const canvasRect = document.getElementById('visual-drop-zone').getBoundingClientRect();
            card.left = e.clientX - canvasRect.left - dragOffsetX;
            card.top = e.clientY - canvasRect.top - dragOffsetY;
            // Bound checks
            if(card.left < 0) card.left = 0;
            if(card.top < 0) card.top = 0;
            
            const cardEl = document.getElementById(card.id);
            cardEl.style.left = card.left + "px";
            cardEl.style.top = card.top + "px";
            renderJoins(); // Redraw lines
        }
    }
}

function addTableToCanvas(tableName, x, y) {
    // Allow duplicate tables? V3 -> No, for simplicity.
    if (tablesOnCanvas.find(t => t.name === tableName)) return;

    const canvasRect = document.getElementById('visual-drop-zone').getBoundingClientRect();
    let left = x ? (x - canvasRect.left) : (tablesOnCanvas.length * 220 + 20);
    let top = y ? (y - canvasRect.top) : 50;
    
    const tableObj = {
        id: "tbl_" + new Date().getTime(),
        name: tableName,
        alias: tableName, // Start alias same as name
        columns: currentData.tables[tableName].schema,
        selected: [],
        left: Math.max(10, left),
        top: Math.max(10, top)
    };
    tablesOnCanvas.push(tableObj);
    renderTableCard(tableObj);
    updateGeneratedSQL();
}

function renderTableCard(tableObj) {
    const canvas = document.getElementById('visual-drop-zone');
    const card = document.createElement('div');
    card.className = 'table-card';
    card.id = tableObj.id;
    card.style.left = tableObj.left + "px";
    card.style.top = tableObj.top + "px";
    card.style.position = "absolute"; 
    
    // FIX: Draggable Header
    const header = document.createElement('div');
    header.className = 'card-header';
    header.draggable = true;
    
    // FIX: Editable Title (Pencil)
    header.innerHTML = `
        <div style="display:flex; align-items:center; gap:5px; flex:1;">
            <span class="title-text">${tableObj.alias}</span>
            <span class="edit-alias" style="font-size:10px; cursor:pointer; opacity:0.7;">✎</span>
        </div>
        <span class="card-close" onclick="removeTable('${tableObj.id}')">✖</span>
    `;
    
    // Alias Edit Logic
    const titleSpan = header.querySelector('.title-text');
    const editBtn = header.querySelector('.edit-alias');
    
    editBtn.onclick = (e) => {
        e.stopPropagation(); // Don't drag
        const input = document.createElement('input');
        input.type = 'text';
        input.value = tableObj.alias;
        input.style.width = "100%";
        input.style.color = "black";
        
        input.onblur = () => {
            tableObj.alias = input.value.trim() || tableObj.name;
            titleSpan.innerText = tableObj.alias;
            input.replaceWith(titleSpan);
            editBtn.style.display = "inline";
            updateGeneratedSQL();
        };
        input.onkeydown = (ev) => { if(ev.key==='Enter') input.blur(); };
        
        editBtn.style.display = "none";
        titleSpan.replaceWith(input);
        input.focus();
    };

    // Drag Start Logic
    header.ondragstart = (e) => {
        const rect = card.getBoundingClientRect();
        dragOffsetX = e.clientX - rect.left;
        dragOffsetY = e.clientY - rect.top;
        e.dataTransfer.setData("application/json", JSON.stringify({ type: "move_card", id: tableObj.id }));
    };

    card.appendChild(header);

    const body = document.createElement('div');
    body.className = 'card-body';

    tableObj.columns.forEach(col => {
        const item = document.createElement('div');
        item.className = 'col-item';
        // Add selected class if already selected
        if(tableObj.selected.includes(col)) {
            item.classList.add('selected');
        }
        item.innerText = col;
        item.dataset.table = tableObj.name; // Refers to source table name, NOT alias (for joins)
        item.dataset.col = col;
        item.draggable = true;

        // FIX: Visual Selection Feedback
        item.onclick = (e) => {
            if(e.target.closest('.col-anchor')) return;
            const idx = tableObj.selected.indexOf(col);
            if(idx > -1) {
                tableObj.selected.splice(idx,1);
                item.classList.remove('selected');
            } else {
                tableObj.selected.push(col);
                item.classList.add('selected');
            }
            updateGeneratedSQL();
        };

        item.ondragstart = (e) => {
            e.stopPropagation();
            e.dataTransfer.setData("application/json", JSON.stringify({
                type: "join_link",
                tableName: tableObj.name, // Use original name for joining logic
                alias: tableObj.alias,
                colName: col
            }));
        };

        item.ondragover = (e) => { e.preventDefault(); item.style.background = "#e6f2ff"; };
        item.ondragleave = (e) => { item.style.background = ""; };
        item.ondrop = (e) => handleJoinDrop(e, tableObj.name, col); // Pass original name

        item.innerHTML += `<div class="col-anchor"></div>`;
        body.appendChild(item);
    });

    card.appendChild(body);
    canvas.appendChild(card);
    
    setTimeout(renderJoins, 0);
}

function removeTable(id) {
    const tbl = tablesOnCanvas.find(t => t.id === id);
    if(tbl) {
        currentData.joins = currentData.joins.filter(j => j.sourceTable !== tbl.name && j.targetTable !== tbl.name);
    }
    tablesOnCanvas = tablesOnCanvas.filter(t => t.id !== id);
    document.getElementById(id).remove();
    renderJoins();
    updateGeneratedSQL();
}

// --- JOIN LOGIC ---
function handleJoinDrop(e, targetTable, targetCol) {
    e.preventDefault();
    e.target.style.background = "";
    
    const raw = e.dataTransfer.getData("application/json");
    if (!raw) return;
    const src = JSON.parse(raw);
    if (src.type !== "join_link") return;
    if (src.tableName === targetTable) return; 

    const newJoin = {
        id: "j_" + new Date().getTime(),
        sourceTable: src.tableName,
        sourceCol: src.colName,
        targetTable: targetTable,
        targetCol: targetCol,
        type: "INNER JOIN",
        op: "="
    };
    currentData.joins.push(newJoin);
    
    renderJoins();
    editJoin(newJoin.id, e.clientX, e.clientY);
    updateGeneratedSQL();
}

function renderJoins() {
    const svg = document.getElementById('connections-layer');
    svg.innerHTML = "";
    const canvasRect = document.getElementById('visual-drop-zone').getBoundingClientRect();

    currentData.joins.forEach(join => {
        // Need to find the Card by TableName first, because there might be multiple instances?
        // For V3 we assume unique table names.
        // We need to look up the DOM element based on dataset
        const srcEl = document.querySelector(`.col-item[data-table="${join.sourceTable}"][data-col="${join.sourceCol}"]`);
        const tgtEl = document.querySelector(`.col-item[data-table="${join.targetTable}"][data-col="${join.targetCol}"]`);

        if (srcEl && tgtEl) {
            const srcRect = srcEl.getBoundingClientRect();
            const tgtRect = tgtEl.getBoundingClientRect();

            const x1 = (srcRect.right - canvasRect.left) - 5;
            const y1 = (srcRect.top - canvasRect.top) + (srcRect.height / 2);
            const x2 = (tgtRect.left - canvasRect.left) + 5;
            const y2 = (tgtRect.top - canvasRect.top) + (tgtRect.height / 2);

            const cp1x = x1 + 50; const cp1y = y1;
            const cp2x = x2 - 50; const cp2y = y2;

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("d", `M ${x1} ${y1} C ${cp1x} ${cp1y}, ${cp2x} ${cp2y}, ${x2} ${y2}`);
            path.setAttribute("class", "join-path");
            if (join.type.includes("LEFT")) path.setAttribute("stroke", "#d13438");
            
            path.onclick = (e) => { e.stopPropagation(); editJoin(join.id, e.clientX, e.clientY); };
            svg.appendChild(path);
        }
    });
}

function editJoin(id, x, y) {
    currentJoinId = id;
    const join = currentData.joins.find(j => j.id === id);
    if(!join) return;

    const pop = document.getElementById('join-editor');
    pop.style.display = "block";
    const popRect = pop.getBoundingClientRect();
    if(x + 260 > window.innerWidth) x -= 260;
    pop.style.left = x + "px";
    pop.style.top = y + "px";

    document.getElementById('join-type-select').value = join.type;
    document.getElementById('join-op-select').value = join.op;
    document.getElementById('join-left-col').innerText = join.sourceCol;
    document.getElementById('join-right-col').innerText = join.targetCol;
}

function saveCurrentJoin() {
    const join = currentData.joins.find(j => j.id === currentJoinId);
    if(join) {
        join.type = document.getElementById('join-type-select').value;
        join.op = document.getElementById('join-op-select').value;
        renderJoins();
        updateGeneratedSQL();
        document.getElementById('join-editor').style.display = "none";
    }
}

function deleteCurrentJoin() {
    currentData.joins = currentData.joins.filter(j => j.id !== currentJoinId);
    renderJoins();
    updateGeneratedSQL();
    document.getElementById('join-editor').style.display = "none";
}

// --- SQL GEN (Updated for Aliases) ---
function updateGeneratedSQL() {
    if(tablesOnCanvas.length === 0) {
        document.getElementById('sql-editor-area').value = "";
        return;
    }

    let selects = [];
    tablesOnCanvas.forEach(t => {
        // Use Alias if present, else Name
        const tblRef = `[${t.alias}]`;
        t.selected.forEach(c => selects.push(`${tblRef}.[${c}]`));
    });

    const mainTable = tablesOnCanvas[0];
    // FROM [RealName] AS [Alias]
    let sql = `SELECT \n    ${selects.length ? selects.join(', ') : '*'} \nFROM [${mainTable.name}] AS [${mainTable.alias}]`;

    // Process Joins
    let joinGroups = {}; 
    let joinOrder = []; 

    currentData.joins.forEach(j => {
        // Find target table object to get its alias
        const targetObj = tablesOnCanvas.find(t => t.name === j.targetTable);
        const sourceObj = tablesOnCanvas.find(t => t.name === j.sourceTable);
        
        if(!targetObj || !sourceObj) return;

        const targetAlias = targetObj.alias;
        const sourceAlias = sourceObj.alias;

        if (!joinGroups[targetAlias]) {
            joinGroups[targetAlias] = {
                type: j.type,
                conditions: [],
                realName: j.targetTable
            };
            joinOrder.push(targetAlias);
        }
        
        joinGroups[targetAlias].conditions.push(
            `[${sourceAlias}].[${j.sourceCol}] ${j.op} [${targetAlias}].[${j.targetCol}]`
        );
    });

    joinOrder.forEach(targetAlias => {
        const grp = joinGroups[targetAlias];
        const combinedConditions = grp.conditions.join(" AND \n    ");
        sql += `\n${grp.type} [${grp.realName}] AS [${targetAlias}] ON ${combinedConditions}`;
    });

    // Orphans
    tablesOnCanvas.slice(1).forEach(t => {
        const isJoined = currentData.joins.some(j => j.targetTable === t.name || j.sourceTable === t.name);
        if (!isJoined) {
            sql += `,\n[${t.name}] AS [${t.alias}]`;
        }
    });

    document.getElementById('sql-editor-area').value = sql;
}

function handleSqlDrop(e) {
    e.preventDefault(); e.stopPropagation();
    document.getElementById('sql-editor-area').style.border = "none";
    let text = e.dataTransfer.getData("text/plain");
    if(!text) return;

    const colName = text.replace(/[\[\]]/g, ''); 
    if (currentData.virtuals[colName]) {
        text = currentData.virtuals[colName]; 
        Object.keys(currentData.virtuals).forEach(key => {
            if (key !== colName && text.includes(`[${key}]`)) {
                text = text.split(`[${key}]`).join(`(${currentData.virtuals[key]})`);
            }
        });
        text = `(${text})`;
    }
    
    const el = document.getElementById('sql-editor-area');
    const start = el.selectionStart;
    el.value = el.value.slice(0, start) + text + el.value.slice(el.selectionEnd);
    el.focus();
}

function openModal(colName) {
    document.getElementById("modal-add-col").style.display = "flex";
    const nameInput = document.getElementById("new-col-name");
    const formulaInput = document.getElementById("new-col-formula");
    const saveBtn = document.getElementById("btn-save-col");

    const modalChipList = document.getElementById("modal-chip-list");
    modalChipList.innerHTML = "";
    
    let allCols = [];
    Object.keys(currentData.tables).forEach(tableName => {
        currentData.tables[tableName].schema.forEach(col => {
            allCols.push({ 
                name: col, 
                type: currentData.tables[tableName].colTypes[col] || "TEXT" 
            });
        });
    });
    
    Object.keys(currentData.virtuals).forEach(v => {
        allCols.push({ name: v, type: "CALC", isVirtual: true });
    });
    
    allCols.forEach(colData => {
        if (colData.name === colName) return; 

        const chip = document.createElement("div");
        chip.className = "chip";
        chip.style.margin = "2px";
        chip.style.transform = "scale(0.9)";
        
        let iconClass = "type-txt", iconText = "ABC";
        if(colData.isVirtual) { iconClass = "type-calc"; iconText = "ƒx"; }
        else if(colData.type === "NUMBER") { iconClass = "type-num"; iconText = "123"; }
        else if(colData.type === "DATE") { iconClass = "type-date"; iconText = "📅"; }

        chip.innerHTML = `
            <div class="chip-icon ${iconClass}" style="width:25px; font-size:10px;">${iconText}</div>
            <div class="chip-name" style="padding:2px 6px;">${colData.name}</div>
        `;
        
        chip.onclick = () => {
            if (activeInput) insertTextAtCursor(activeInput, `[${colData.name}]`);
            else insertTextAtCursor(formulaInput, `[${colData.name}]`);
        };
        modalChipList.appendChild(chip);
    });

    document.getElementById("cond-rows-container").innerHTML = "";
    document.getElementById("cond-else").value = "";
    document.getElementById("cond-else").onfocus = (e) => activeInput = e.target;

    if (colName) {
        currentEditCol = colName;
        nameInput.value = colName;
        nameInput.disabled = true;
        
        if (currentData.virtuals[colName] && currentData.virtuals[colName].trim().toUpperCase().startsWith("CASE")) {
             formulaInput.value = currentData.virtuals[colName];
             switchModalTab('formula');
        } else {
             formulaInput.value = currentData.virtuals[colName] || "";
             switchModalTab('formula');
        }
        saveBtn.innerText = "Update";
    } else {
        currentEditCol = null;
        nameInput.value = "";
        nameInput.disabled = false;
        formulaInput.value = "";
        saveBtn.innerText = "Add";
        switchModalTab('formula');
        addConditionRow(); 
    }
    
    formulaInput.onfocus = (e) => activeInput = e.target;
    formulaInput.focus();
    activeInput = formulaInput;
}

function closeModal() {
    document.getElementById("modal-add-col").style.display = "none";
    activeInput = null;
}

function saveVirtualColumn() {
    const name = document.getElementById("new-col-name").value.trim();
    let formula = "";

    if (document.getElementById("tab-btn-cond").classList.contains("active")) {
        formula = buildCaseStatement();
    } else {
        formula = document.getElementById("new-col-formula").value.trim();
    }

    if (!name || !formula) return;

    currentData.virtuals[name] = formula;
    renderIngredients();
    closeModal();
}

function addConditionRow() {
    const container = document.getElementById("cond-rows-container");
    const div = document.createElement("div");
    div.className = "cond-row";
    
    div.innerHTML = `
        <span class="cond-if">IF</span>
        <input type="text" class="cond-input-col" placeholder="Column">
        <select class="cond-op"><option>=</option><option>></option><option><</option><option>>=</option><option><=</option><option>!=</option><option>LIKE</option></select>
        <input type="text" class="cond-val" placeholder="Value">
        <span class="cond-then">THEN</span>
        <input type="text" class="cond-res" placeholder="Result">
        <span class="cond-del">✖</span>
    `;
    
    div.querySelectorAll("input").forEach(inp => inp.onfocus = (e) => activeInput = e.target);
    div.querySelector(".cond-del").onclick = () => div.remove();
    container.appendChild(div);
    
    const first = div.querySelector(".cond-input-col");
    if(first) { first.focus(); activeInput = first; }
}

function buildCaseStatement() {
    const rows = document.querySelectorAll(".cond-row");
    let sql = "CASE ";
    rows.forEach(row => {
        let col = row.querySelector(".cond-input-col").value.trim();
        let op = row.querySelector(".cond-op").value;
        let val = row.querySelector(".cond-val").value.trim();
        let res = row.querySelector(".cond-res").value.trim();
        if(col) sql += `WHEN ${col} ${op} ${val} THEN ${res} `;
    });
    sql += `ELSE ${document.getElementById("cond-else").value || 'NULL'} END`;
    return sql;
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

function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.getElementById('tab-' + tabName).classList.add('active');
    document.getElementById('view-visual').style.display = 'none';
    document.getElementById('view-sql').style.display = 'none';
    document.getElementById('view-' + tabName).style.display = 'flex';
}

function switchModalTab(tabName) {
    document.getElementById("tab-btn-formula").classList.remove("active");
    document.getElementById("tab-btn-cond").classList.remove("active");
    document.getElementById("view-formula").style.display = "none";
    document.getElementById("view-cond").style.display = "none";

    document.getElementById("tab-btn-" + tabName).classList.add("active");
    document.getElementById("view-" + tabName).style.display = "block";
    
    if (tabName === 'formula') {
        const formulaInput = document.getElementById("new-col-formula");
        formulaInput.focus();
        activeInput = formulaInput;
    } else {
        activeInput = null;
    }
}

window.updateJoinType = (id, type) => {
    const join = currentData.joins.find(j => j.id === id);
    if(join) { join.type = type; updateGeneratedSQL(); }
};
window.removeJoin = (id) => {
    currentData.joins = currentData.joins.filter(j => j.id !== id);
    renderJoins();
    updateGeneratedSQL();
};
window.removeTable = (id) => {
    // Defined in file but exposed here
    // Already defined in function scope above, good to go
    const tbl = tablesOnCanvas.find(t => t.id === id);
    if(tbl) {
        currentData.joins = currentData.joins.filter(j => j.sourceTable !== tbl.name && j.targetTable !== tbl.name);
    }
    tablesOnCanvas = tablesOnCanvas.filter(t => t.id !== id);
    document.getElementById(id).remove();
    renderJoins();
    updateGeneratedSQL();
};
