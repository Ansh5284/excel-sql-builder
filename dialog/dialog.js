/* global Office, document, console */

// Local state
let currentData = {
    tables: {}, 
    ctes: {}, // { "Name": { sql: "", schema: [], visualState: {}, unionConfig: {}, baseTables: [], dependencies: [] } }
    virtuals: {},
    joins: [] 
};

// UNION BUILDER STATE
let unionState = {
    master: null,
    others: [] 
};

// LOGIC DECK STATE
let currentLogicState = {
    WHERE: [],  // Array of Group/Rule objects
    HAVING: []
};
let activeLogicTab = 'WHERE'; // 'WHERE' or 'HAVING'

let tablesOnCanvas = []; 
let activeInput = null; 
let currentEditCol = null; 
let currentJoinId = null;
let editingCTEName = null; 
let confirmCallback = null; 

// Drag State
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
        
        // CTE Listeners
        document.getElementById('btn-save-view-cte').onclick = openNameCTEModal;
        document.getElementById('btn-close-cte-name').onclick = closeNameCTEModal;
        document.getElementById('btn-confirm-save-cte').onclick = saveViewAsCTE;

        // UNION Listeners
        document.getElementById('btn-create-union').onclick = () => { editingCTEName = null; openUnionModal(); };
        document.getElementById('btn-close-union').onclick = closeUnionModal;
        document.getElementById('btn-cancel-union').onclick = closeUnionModal;
        document.getElementById('btn-save-union').onclick = saveUnionAsCTE;
        document.getElementById('union-master-select').onchange = updateUnionMaster;

        document.getElementById('btn-cancel-col').onclick = closeModal;
        document.getElementById('btn-close-x').onclick = closeModal;
        document.getElementById('btn-save-col').onclick = saveVirtualColumn;
        document.getElementById('btn-add-rule').onclick = addConditionRow;

        document.getElementById('btn-delete-join').onclick = deleteCurrentJoin;
        document.getElementById('btn-save-join').onclick = saveCurrentJoin;
        
        // Confirm Modal Listeners
        document.getElementById('btn-close-confirm').onclick = closeConfirmModal;
        document.getElementById('btn-cancel-confirm').onclick = closeConfirmModal;
        document.getElementById('btn-confirm-yes').onclick = () => {
            if (confirmCallback) confirmCallback();
            closeConfirmModal();
        };

        // Logic Deck Listeners
        window.switchLogicTab = switchLogicTab;
        window.toggleDeck = toggleDeck;
        // Init Logic Deck Drop
        const logicZone = document.getElementById('logicRootZone');
        logicZone.ondragover = handleLogicDragOver;
        logicZone.ondragleave = handleLogicDragLeave;
        logicZone.ondrop = (e) => handleLogicDrop(e, null);

        // Apply Button (Run)
        document.getElementById('btn-apply').onclick = handleApply;

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

// --- LOGIC DECK LOGIC ---

function switchLogicTab(tabName) {
    activeLogicTab = tabName;
    document.getElementById('tab-logic-where').classList.remove('active');
    document.getElementById('tab-logic-having').classList.remove('active');
    document.getElementById(`tab-logic-${tabName.toLowerCase()}`).classList.add('active');
    renderLogicDeck();
}

function toggleDeck() {
    const d = document.getElementById('logicDeck');
    if (d.style.height === "35px") d.style.height = "350px";
    else d.style.height = "35px";
}

function handleLogicDragOver(ev) {
    ev.preventDefault();
    ev.stopPropagation();
    
    const targetGroup = ev.target.closest('.logic-group');
    const targetRoot = ev.target.closest('.deck-body');

    if (targetGroup) {
        targetGroup.classList.add('drag-over-group');
        if(targetRoot) targetRoot.classList.remove('drag-over');
    } else if (targetRoot) {
        targetRoot.classList.add('drag-over');
    }
}

function handleLogicDragLeave(ev) {
    if(ev.target.classList.contains('logic-group')) {
        ev.target.classList.remove('drag-over-group');
    }
    if(ev.target.id === 'logicRootZone') {
        ev.target.classList.remove('drag-over');
    }
}

function handleLogicDrop(ev, parentGroup) {
    ev.preventDefault();
    ev.stopPropagation();
    
    document.querySelectorAll('.drag-over, .drag-over-group').forEach(el => {
        el.classList.remove('drag-over');
        el.classList.remove('drag-over-group');
    });

    const raw = ev.dataTransfer.getData("application/json");
    if(!raw) return;
    const payload = JSON.parse(raw); // { type, tableName, colName, alias } OR { type: 'table', name: 'CTE' }

    // Convert payload to Logic Rule Format
    let newRule = {
        type: 'rule',
        left: null,
        op: '=',
        right: { type: 'value', val: '' }
    };

    if (payload.type === 'join_link') {
        // Dragged a column
        const alias = payload.alias || payload.tableName;
        newRule.left = { type: 'col', val: `[${alias}].[${payload.colName}]`, label: payload.colName };
    } else if (payload.type === 'table') {
        // Dragged a CTE (assuming we allow dropping CTE as left operand? Probably rare, usually right)
        // Let's assume user dropped it onto root to start a rule
        if (currentData.ctes[payload.name]) {
             // CTEs usually go on Right side (IN (Select...)) but here we initiate rule
             // Let's default left to empty and let them drop? Or create a rule where CTE is right?
             // Prototype logic allowed dragging Column to start. 
             // If CTE dropped, maybe we can't start a rule easily without knowing col.
             // Let's just return if not a column for now as root drop
             return;
        }
    } else {
        return;
    }

    // Determine Insertion Point
    const dropTargetGroup = ev.target.closest('.logic-group');
    const targetArray = currentLogicState[activeLogicTab];

    if (dropTargetGroup) {
        // Dropped INSIDE a group -> Add to that group
        // In real app we need to map DOM to Data. For this impl, we'll traverse via ID or re-render.
        // Simplified: We need a way to link DOM group to data object.
        // Let's use array index mapping or simple ID.
        // Adding ID to data objects:
        const groupId = dropTargetGroup.dataset.id;
        const groupObj = findGroupById(targetArray, groupId);
        if(groupObj) {
            groupObj.children.push(newRule);
        }
    } else {
        // Dropped on Root
        if (targetArray.length === 0) {
            // First item
            targetArray.push(newRule);
        } else {
            // Wrap in new Parent Group
            const newGroup = {
                id: Date.now().toString(),
                type: 'group',
                op: 'AND', // Default outer join
                children: [...targetArray, newRule] // Move existing + new
            };
            currentLogicState[activeLogicTab] = [newGroup]; // Replace root with this group
        }
    }
    renderLogicDeck();
    updateGeneratedSQL();
}

function findGroupById(arr, id) {
    for(let item of arr) {
        if(item.type === 'group') {
            if(item.id === id) return item;
            const found = findGroupById(item.children, id);
            if(found) return found;
        }
    }
    return null;
}

// --- RENDERING LOGIC DECK ---

function renderLogicDeck() {
    const root = document.getElementById('logicRootZone');
    const emptyMsg = document.getElementById('logicEmptyMsg');
    
    // Clear content except empty msg
    Array.from(root.children).forEach(c => { if(c.id !== 'logicEmptyMsg') c.remove(); });

    const items = currentLogicState[activeLogicTab];
    if (items.length === 0) {
        emptyMsg.style.display = 'block';
    } else {
        emptyMsg.style.display = 'none';
        items.forEach(item => {
            if(item.type === 'group') {
                root.appendChild(createGroupDOM(item));
            } else {
                root.appendChild(createRuleDOM(item, items)); // Pass array for delete
            }
        });
    }
}

function createGroupDOM(groupData) {
    const div = document.createElement('div');
    div.className = 'logic-group';
    div.dataset.id = groupData.id;
    
    div.ondragover = handleLogicDragOver;
    div.ondragleave = handleLogicDragLeave;
    div.ondrop = handleLogicDrop;

    // Add children
    groupData.children.forEach((child, index) => {
        // Connector (if not first)
        if (index > 0) {
            const conn = document.createElement('div');
            conn.className = 'logic-connector';
            conn.innerText = groupData.op;
            conn.dataset.op = groupData.op;
            conn.onclick = (e) => {
                e.stopPropagation();
                groupData.op = groupData.op === "AND" ? "OR" : "AND";
                renderLogicDeck();
                updateGeneratedSQL();
            };
            div.appendChild(conn);
        }

        if (child.type === 'group') {
            div.appendChild(createGroupDOM(child));
        } else {
            div.appendChild(createRuleDOM(child, groupData.children));
        }
    });

    return div;
}

function createRuleDOM(rule, parentArray) {
    const row = document.createElement('div');
    row.className = 'logic-row';

    // 1. Left (Column)
    if (rule.left && rule.left.type === 'col') {
        row.innerHTML = `<span class="col-badge">${rule.left.label}</span>`;
    }

    // 2. Operator
    const opSelect = document.createElement('select');
    opSelect.className = "op-select";
    ["=", ">", "<", ">=", "<=", "<>", "IN", "NOT IN", "LIKE"].forEach(o => {
        const opt = document.createElement('option');
        opt.value = o;
        opt.innerText = o;
        if(o === rule.op) opt.selected = true;
        opSelect.appendChild(opt);
    });
    opSelect.onchange = () => { rule.op = opSelect.value; updateGeneratedSQL(); };
    row.appendChild(opSelect);

    // 3. Right (Input/CTE)
    const rightContainer = document.createElement('div');
    // Decide what to show
    if (rule.right.type === 'cte') {
        const badge = document.createElement('span');
        badge.className = 'cte-badge';
        badge.innerText = `⚡ ${rule.right.label}`;
        badge.onclick = () => {
            // Remove CTE, revert to input
            rule.right = { type: 'value', val: '' };
            renderLogicDeck();
            updateGeneratedSQL();
        };
        rightContainer.appendChild(badge);
    } else {
        const input = document.createElement('input');
        input.className = "value-input";
        input.value = rule.right.val;
        input.placeholder = "Value or CTE";
        input.oninput = (e) => { rule.right.val = e.target.value; updateGeneratedSQL(); };
        
        // Handle CTE Drop on Input
        input.ondragover = (e) => { e.preventDefault(); input.classList.add('drag-over'); };
        input.ondragleave = () => input.classList.remove('drag-over');
        input.ondrop = (e) => handleValueDrop(e, rule);

        rightContainer.appendChild(input);
    }
    row.appendChild(rightContainer);

    // 4. Delete
    const del = document.createElement('span');
    del.className = "btn-del-rule";
    del.innerHTML = "&times;";
    del.onclick = () => {
        const idx = parentArray.indexOf(rule);
        if (idx > -1) parentArray.splice(idx, 1);
        
        // Clean up empty groups logic (if needed, or re-render handles it)
        // If root array becomes empty, it's fine.
        // If a group becomes empty or has 1 child, we might want to flatten.
        // For MVP, just re-render.
        renderLogicDeck();
        updateGeneratedSQL();
    };
    row.appendChild(del);

    return row;
}

function handleValueDrop(ev, rule) {
    ev.preventDefault();
    ev.stopPropagation();
    
    const raw = ev.dataTransfer.getData("application/json");
    if(!raw) return;
    const payload = JSON.parse(raw);

    // FIX: Allow both CTEs AND Regular Tables
    let schema = [];
    if (payload.type === 'table') {
        if (currentData.ctes[payload.name]) {
            schema = currentData.ctes[payload.name].schema;
        } else if (currentData.tables[payload.name]) {
            schema = currentData.tables[payload.name].schema;
        }
        
        if (schema.length > 0) {
            // Valid table source
            if (schema.length === 1) {
                // Auto Select Single Col
                rule.right = { 
                    type: 'cte', 
                    val: `(SELECT [${schema[0]}] FROM [${payload.name}])`, 
                    label: `${payload.name}.${schema[0]}`,
                    tableName: payload.name // FIX: Store table name for dependency check
                };
                renderLogicDeck();
                updateGeneratedSQL();
            } else {
                showCteColSelector(payload.name, schema, rule);
            }
        }
    }
}

function showCteColSelector(cteName, cols, ruleObj) {
    const modal = document.getElementById('cteModal');
    document.getElementById('modalCteName').innerText = cteName;
    const list = document.getElementById('modalList');
    list.innerHTML = "";

    cols.forEach(col => {
        const item = document.createElement('div');
        item.className = 'popover-item';
        item.innerText = col;
        item.onclick = () => {
            ruleObj.right = { type: 'cte', val: `(SELECT [${col}] FROM [${cteName}])`, label: `${cteName}.${col}` };
            modal.style.display = 'none';
            renderLogicDeck();
            updateGeneratedSQL();
        };
        list.appendChild(item);
    });

    modal.style.display = 'block';
}

// --- UPDATE SQL GENERATOR FOR LOGIC DECK ---

function buildLogicString(items) {
    if (items.length === 0) return "";
    
    // items is array of rules/groups.
    // If multiple items at root level, what connects them?
    // Our data structure: [ {type: 'group', op: 'AND', children: [rule1, rule2]} ]
    // The visual builder usually ensures Root is either a list of rules (implicit AND? No, we added manual connector logic).
    // Actually, our Add Drop Logic wraps things in a Group if collision.
    // So usually root has 1 item (Group) or 1 item (Rule).
    // If flat list [Rule1, Rule2], we need implicit connector.
    // Let's assume Root Level is Implicit AND if multiple.
    
    return items.map(item => parseItem(item)).join(" AND ");
}

function parseItem(item) {
    if (item.type === 'group') {
        const childrenSQL = item.children.map(c => parseItem(c));
        return `(${childrenSQL.join(` ${item.op} `)})`;
    } else {
        // Rule
        const rightVal = item.right.type === 'cte' ? item.right.val : `'${item.right.val}'`; // Add quotes for text?
        // Basic type check for quotes? For MVP assume string literal unless number.
        // Better: Check if number.
        let val = item.right.val;
        if(item.right.type !== 'cte' && isNaN(val) && !val.startsWith("'")) {
             val = `'${val}'`;
        }
        
        return `${item.left.val} ${item.op} ${val}`;
    }
}

// --- UNION BUILDER LOGIC (Existing...) ---
function openUnionModal() {
    // If NOT editing, reset state
    if (!editingCTEName) {
        unionState = { master: null, others: [] };
        document.getElementById('union-master-cols').innerHTML = "";
        document.getElementById('union-mapping-container').innerHTML = "";
        document.getElementById('union-step-mapping').style.display = "none";
        document.getElementById('union-step-name').style.display = "none";
        document.getElementById('union-result-name').value = "";
        document.getElementById('btn-save-union').innerText = "Save as Union CTE";
    } else {
        // Editing Mode: State is already prepared by loadCTEForEdit
        document.getElementById('btn-save-union').innerText = "Update Union CTE";
    }

    // Populate Sidebar
    const sidebarList = document.getElementById('union-source-list');
    sidebarList.innerHTML = "";
    Object.keys(currentData.tables).forEach(name => {
        const item = document.createElement('div');
        item.className = 'table-item';
        item.innerHTML = `${name} <span class="add-icon">+</span>`;
        item.onclick = () => addTableToUnion(name);
        sidebarList.appendChild(item);
    });

    // Populate Master Select
    const masterSelect = document.getElementById('union-master-select');
    masterSelect.innerHTML = '<option value="" disabled selected>Select a primary table...</option>';
    Object.keys(currentData.tables).forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.innerText = name;
        if (unionState.master === name) opt.selected = true;
        masterSelect.appendChild(opt);
    });

    document.getElementById('modal-union-builder').style.display = 'flex';

    // If Editing, restore UI state
    if (editingCTEName && unionState.master) {
        // Trigger column display
        const cols = currentData.tables[unionState.master].schema;
        const container = document.getElementById('union-master-cols');
        container.innerHTML = "Output Columns: " + cols.map(c => `<span class='chip' style='font-size:10px; padding:2px 6px; margin:2px;'>${c}</span>`).join("");
        
        document.getElementById('union-step-mapping').style.display = 'block';
        document.getElementById('union-step-name').style.display = 'block';
        
        // Restore Name
        document.getElementById('union-result-name').value = editingCTEName;
        
        // Restore Radio
        const savedConfig = currentData.ctes[editingCTEName].unionConfig;
        if (savedConfig && savedConfig.type) {
             const radios = document.getElementsByName("u_type");
             radios.forEach(r => { if(r.value === savedConfig.type) r.checked = true; });
        }

        // Render Grid & Restore Mappings
        renderUnionMapping(savedConfig ? savedConfig.mappings : null);
    }
}

function closeUnionModal() {
    document.getElementById('modal-union-builder').style.display = 'none';
    editingCTEName = null; // Clear edit state on close
}

function updateUnionMaster() {
    const val = document.getElementById('union-master-select').value;
    if (!val) return;
    
    unionState.master = val;
    // Remove master from others if present
    unionState.others = unionState.others.filter(t => t.name !== val);

    // Show columns
    const cols = currentData.tables[val].schema;
    const container = document.getElementById('union-master-cols');
    container.innerHTML = "Output Columns: " + cols.map(c => `<span class='chip' style='font-size:10px; padding:2px 6px; margin:2px;'>${c}</span>`).join("");

    document.getElementById('union-step-mapping').style.display = 'block';
    document.getElementById('union-step-name').style.display = 'block';
    renderUnionMapping();
}

function addTableToUnion(tableName) {
    if (!unionState.master) {
        alert("Please select a Master Table (Step 1) first.");
        return;
    }
    if (unionState.master === tableName) return;
    if (unionState.others.find(t => t.name === tableName)) return;

    unionState.others.push({ name: tableName });
    renderUnionMapping();
}

function removeTableFromUnion(name) {
    unionState.others = unionState.others.filter(t => t.name !== name);
    renderUnionMapping();
}

function renderUnionMapping(restoredMappings = null) {
    const container = document.getElementById('union-mapping-container');
    container.innerHTML = "";

    if (unionState.others.length === 0) {
        container.innerHTML = "<div style='color:#999; font-style:italic; padding:10px;'>Add tables from the left sidebar to start mapping.</div>";
        return;
    }

    const masterCols = currentData.tables[unionState.master].schema;

    // Table Structure
    let html = `<table class="mapping-table"><thead><tr><th style="width:150px;">Target (${unionState.master})</th>`;
    unionState.others.forEach(t => {
        html += `<th>${t.name} <span onclick="removeTableFromUnion('${t.name}')" style="color:red; cursor:pointer; font-weight:bold; margin-left:5px;">(x)</span></th>`;
    });
    html += `</tr></thead><tbody>`;

    // Rows
    masterCols.forEach(mCol => {
        html += `<tr><td>${mCol}</td>`;
        
        unionState.others.forEach(otherTbl => {
            const otherCols = currentData.tables[otherTbl.name].schema;
            
            // Check if we have a saved mapping for this cell
            let savedVal = null;
            if (restoredMappings && restoredMappings[otherTbl.name] && restoredMappings[otherTbl.name][mCol]) {
                savedVal = restoredMappings[otherTbl.name][mCol];
            }

            // Fuzzy Match Logic (Only if no saved value)
            let selected = savedVal;
            
            if (!selected) {
                const mLow = mCol.toLowerCase().replace(/[^a-z0-9]/g, "");
                
                // 1. Exact Match
                if (otherCols.includes(mCol)) selected = mCol;
                else {
                    // 2. Fuzzy Containment
                    for (let oc of otherCols) {
                        const oLow = oc.toLowerCase().replace(/[^a-z0-9]/g, "");
                        if (oLow === mLow || oLow.includes(mLow) || mLow.includes(oLow)) {
                            selected = oc;
                            break;
                        }
                    }
                }
            }

            let options = `<option value='NULL'>-- NULL --</option>`;
            otherCols.forEach(oc => {
                const isSel = (oc === selected) ? "selected" : "";
                options += `<option value='${oc}' ${isSel}>${oc}</option>`;
            });

            // Set explicit NULL selection if saved was NULL
            if (selected === 'NULL') {
                 options = options.replace("value='NULL'", "value='NULL' selected");
            }

            html += `<td><select class="map-select" data-tbl="${otherTbl.name}" data-target="${mCol}">${options}</select></td>`;
        });
        html += `</tr>`;
    });

    html += `</tbody></table>`;
    container.innerHTML = html;

    // Re-attach removal handlers via global (simplified for add-in context)
    window.removeTableFromUnion = removeTableFromUnion;
}

function saveUnionAsCTE() {
    if (!unionState.master) return;
    const name = document.getElementById('union-result-name').value.trim();
    if (!name) {
        alert("Please name your Union CTE.");
        return;
    }

    const unionType = document.querySelector('input[name="u_type"]:checked').value;
    const masterCols = currentData.tables[unionState.master].schema;
    
    // 0. CAPTURE MAPPINGS (For Edit Restoration)
    const currentMappings = {};

    // 1. GENERATE SQL
    let sql = `SELECT ${masterCols.map(c => `[${c}]`).join(", ")}\nFROM [${unionState.master}]`;

    unionState.others.forEach(otherTbl => {
        sql += `\n\n${unionType}\n\nSELECT `;
        currentMappings[otherTbl.name] = {};
        
        let cols = [];
        masterCols.forEach(mCol => {
            const sel = document.querySelector(`.map-select[data-tbl="${otherTbl.name}"][data-target="${mCol}"]`);
            if (sel) {
                const val = sel.value;
                currentMappings[otherTbl.name][mCol] = val; // Store for config
                
                if (val === 'NULL') cols.push(`NULL AS [${mCol}]`);
                else cols.push(`[${val}] AS [${mCol}]`);
            } else {
                cols.push(`NULL AS [${mCol}]`);
                currentMappings[otherTbl.name][mCol] = 'NULL';
            }
        });
        sql += `${cols.join(", ")}\nFROM [${otherTbl.name}]`;
    });

    // 2. IDENTIFY DEPENDENCIES
    let baseTables = [];
    let cteDependencies = [];
    const allTables = [unionState.master, ...unionState.others.map(t => t.name)];

    allTables.forEach(tName => {
        if (currentData.ctes[tName]) {
            cteDependencies.push(tName);
        } else {
            baseTables.push(tName);
        }
    });

    // 3. SAVE CTE with UnionConfig
    const schemaMap = {};
    masterCols.forEach(c => schemaMap[c] = currentData.tables[unionState.master].colTypes[c] || "TEXT");

    currentData.ctes[name] = {
        sql: sql,
        schema: masterCols,
        colTypes: schemaMap,
        baseTables: [...new Set(baseTables)],
        dependencies: [...new Set(cteDependencies)],
        visualState: null, // Unions are not editable in visual joiner
        unionConfig: {
            master: unionState.master,
            others: JSON.parse(JSON.stringify(unionState.others)),
            type: unionType,
            mappings: currentMappings
        }
    };

    // 4. MOCK AS TABLE
    currentData.tables[name] = {
        schema: masterCols,
        colTypes: schemaMap
    };

    renderIngredients();
    closeUnionModal();
}


// --- CONFIRMATION MODAL HELPERS ---
function showConfirm(title, msg, cb) {
    document.getElementById('confirm-title').innerText = title;
    document.getElementById('confirm-message').innerText = msg;
    confirmCallback = cb;
    document.getElementById('modal-confirm').style.display = 'flex';
}

function closeConfirmModal() {
    document.getElementById('modal-confirm').style.display = 'none';
    confirmCallback = null;
}

// --- CTE LOGIC: SAVE VIEW ---

function openNameCTEModal() {
    if (tablesOnCanvas.length === 0) {
        // No visual elements
        return; 
    }
    document.getElementById("modal-name-cte").style.display = "flex";
    const nameInput = document.getElementById("input-cte-name");
    
    if (editingCTEName) {
        nameInput.value = editingCTEName;
        document.getElementById("btn-confirm-save-cte").innerText = "Update CTE & Clear";
    } else {
        nameInput.value = "";
        document.getElementById("btn-confirm-save-cte").innerText = "Save & Clear Canvas";
    }
    nameInput.focus();
}

function closeNameCTEModal() {
    document.getElementById("modal-name-cte").style.display = "none";
}

function saveViewAsCTE() {
    const name = document.getElementById("input-cte-name").value.trim();
    if (!name) return;

    // 1. INFER SCHEMA (Output Columns)
    // We look at what columns the user has "Checked" (selected) on the tables.
    let outputSchema = [];
    let schemaMap = {}; // "Alias": Type
    
    tablesOnCanvas.forEach(t => {
        t.selected.forEach(col => {
            // Did user alias it?
            const finalName = (t.columnAliases && t.columnAliases[col]) ? t.columnAliases[col] : col;
            outputSchema.push(finalName);
            
            // Infer type
            let type = "TEXT";
            if (currentData.tables[t.name]) {
                type = currentData.tables[t.name].colTypes[col] || "TEXT";
            }
            schemaMap[finalName] = type;
        });
    });

    if (outputSchema.length === 0) {
        // Fallback: If nothing selected, select ALL from first table? 
        // Or just alert user.
        // For MVP, let's assume valid selection.
    }

    // 2. GENERATE INNER SQL
    // We reuse the existing generation logic, but we capture the RESULT, not set the textarea.
    const innerSQL = generateSQLFromCanvas();

    // 3. IDENTIFY DEPENDENCIES
    // Which real tables (Excel) and which other CTEs are used here?
    let baseTables = [];
    let cteDependencies = [];

    tablesOnCanvas.forEach(t => {
        if (currentData.ctes[t.name]) {
            // It uses another CTE
            cteDependencies.push(t.name);
        } else {
            // It uses a base table
            baseTables.push(t.name);
        }
    });

    // Check Logic Deck dependencies too (CTEs used in Subqueries)
    // currentLogicState.WHERE... check right.type === 'cte'
    // Simplified recursive check needed if deep.
    // For now assuming visual tables cover dependencies or user manually adds them.
    // Robust way: Traverse logic state
    function traverseLogicForDeps(arr) {
        arr.forEach(item => {
            if(item.type === 'group') traverseLogicForDeps(item.children);
            else if(item.right.type === 'cte') {
                // item.right.label is "CTEName.ColName" or "CTEName"
                const cteName = item.right.label.split('.')[0];
                if(currentData.ctes[cteName]) cteDependencies.push(cteName);
            }
        });
    }
    traverseLogicForDeps(currentLogicState.WHERE);
    traverseLogicForDeps(currentLogicState.HAVING);


    baseTables = [...new Set(baseTables)];
    cteDependencies = [...new Set(cteDependencies)];

    // 4. CAPTURE VISUAL STATE
    // We must deep clone the state so subsequent canvas edits don't mutate the saved CTE
    const visualState = {
        tables: JSON.parse(JSON.stringify(tablesOnCanvas)),
        joins: JSON.parse(JSON.stringify(currentData.joins)),
        logic: JSON.parse(JSON.stringify(currentLogicState)) // NEW: Save Logic
    };

    // 5. SAVE CTE DEFINITION
    currentData.ctes[name] = {
        sql: innerSQL,
        schema: outputSchema,
        colTypes: schemaMap,
        visualState: visualState,
        baseTables: baseTables,
        dependencies: cteDependencies
    };

    // 6. MOCK AS TABLE FOR UI
    currentData.tables[name] = {
        schema: outputSchema,
        colTypes: schemaMap // Inherited types
    };

    // 7. CLEAN UP
    closeNameCTEModal();
    clearCanvas();
    editingCTEName = null;
    renderIngredients();
    
    // Clear SQL Editor
    document.getElementById("sql-editor-area").value = "";
}

function loadCTEForEdit(name) {
    const cte = currentData.ctes[name];
    if (!cte) return;
    
    // CHECK TYPE: UNION or VIEW?
    if (cte.unionConfig) {
        // It is a Union CTE
        editingCTEName = name;
        unionState = {
            master: cte.unionConfig.master,
            others: JSON.parse(JSON.stringify(cte.unionConfig.others))
        };
        openUnionModal(); // Logic inside openUnionModal handles the hydration
        return;
    }

    // Otherwise, it's a Visual View CTE
    if (!cte.visualState) {
        alert("Unknown CTE type. Cannot edit.");
        return;
    }

    // RESTORE STATE
    tablesOnCanvas = JSON.parse(JSON.stringify(cte.visualState.tables));
    currentData.joins = JSON.parse(JSON.stringify(cte.visualState.joins));
    // Restore Logic if exists
    if (cte.visualState.logic) {
        currentLogicState = JSON.parse(JSON.stringify(cte.visualState.logic));
    } else {
        currentLogicState = { WHERE: [], HAVING: [] };
    }
    
    editingCTEName = name;

    // REDRAW
    const canvas = document.getElementById('visual-drop-zone');
    const svg = document.getElementById('connections-layer');
    canvas.innerHTML = '';
    canvas.appendChild(svg);

    tablesOnCanvas.forEach(t => renderTableCard(t));
    renderJoins();
    renderLogicDeck(); // NEW
    updateGeneratedSQL();
}

function deleteCTE(name) {
    if (currentData.ctes[name]) {
        delete currentData.ctes[name];
    }
    if (currentData.tables[name]) {
        delete currentData.tables[name];
    }
    // Also remove any active editing state if we deleted what we were editing
    if (editingCTEName === name) {
        editingCTEName = null;
    }
    
    renderIngredients();
}

function clearCanvas() {
    tablesOnCanvas = [];
    currentData.joins = [];
    currentLogicState = { WHERE: [], HAVING: [] }; // Reset Logic
    
    const canvas = document.getElementById('visual-drop-zone');
    const svg = document.getElementById('connections-layer');
    canvas.innerHTML = '';
    canvas.appendChild(svg);
    // Re-add empty state text
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.innerText = 'Drag Tables here, connect them, then click "Save View as CTE" to build logic steps.';
    canvas.appendChild(empty);
    
    renderJoins(); // Clears lines
    renderLogicDeck(); // Clears deck
}

// --- HELPER: Calculate Global Sort Rank for Badge Display ---
function getGlobalSortRank(tableId, colName) {
    // 1. Collect ALL sorts from ALL tables
    let allSorts = [];
    tablesOnCanvas.forEach(t => {
        if(t.sorts) {
            Object.keys(t.sorts).forEach(c => {
                allSorts.push({
                    tId: t.id,
                    col: c,
                    ts: t.sorts[c].ts // Timestamp/Index
                });
            });
        }
    });

    // 2. Sort by timestamp (asc) to get hierarchy 1, 2, 3...
    allSorts.sort((a,b) => a.ts - b.ts);

    // 3. Find index of requested column
    const idx = allSorts.findIndex(x => x.tId === tableId && x.col === colName);
    return idx + 1; // 1-based rank
}

// --- SQL GENERATION & EXECUTION ---

function generateSQLFromCanvas() {
    // Reused logic from updateGeneratedSQL but returns string instead of setting UI
    if(tablesOnCanvas.length === 0) return "";

    let selects = [];
    let orderBys = []; // NEW: Collection for sort

    tablesOnCanvas.forEach(t => {
        const tblRef = `[${t.alias}]`;
        t.selected.forEach(c => {
            const colAlias = t.columnAliases && t.columnAliases[c];
            if (colAlias) {
                selects.push(`${tblRef}.[${c}] AS [${colAlias}]`);
            } else {
                selects.push(`${tblRef}.[${c}]`);
            }
        });

        // NEW: Collect Sorts
        if(t.sorts) {
            Object.keys(t.sorts).forEach(c => {
                orderBys.push({
                    val: `${tblRef}.[${c}] ${t.sorts[c].dir}`,
                    ts: t.sorts[c].ts
                });
            });
        }
    });

    const mainTable = tablesOnCanvas[0];
    let sql = `SELECT \n    ${selects.length ? selects.join(', ') : '*'} \nFROM [${mainTable.name}] AS [${mainTable.alias}]`;

    let joinGroups = {}; 
    let joinOrder = []; 

    currentData.joins.forEach(j => {
        const targetObj = tablesOnCanvas.find(t => t.name === j.targetTable);
        const sourceObj = tablesOnCanvas.find(t => t.name === j.sourceTable);
        if(!targetObj || !sourceObj) return;

        const targetAlias = targetObj.alias;
        const sourceAlias = sourceObj.alias;

        if (!joinGroups[targetAlias]) {
            joinGroups[targetAlias] = { type: j.type, conditions: [], realName: j.targetTable };
            joinOrder.push(targetAlias);
        }
        joinGroups[targetAlias].conditions.push(`[${sourceAlias}].[${j.sourceCol}] ${j.op} [${targetAlias}].[${j.targetCol}]`);
    });

    joinOrder.forEach(targetAlias => {
        const grp = joinGroups[targetAlias];
        sql += `\n${grp.type} [${grp.realName}] AS [${targetAlias}] ON ${grp.conditions.join(" AND \n    ")}`;
    });

    tablesOnCanvas.slice(1).forEach(t => {
        const isJoined = currentData.joins.some(j => j.targetTable === t.name || j.sourceTable === t.name);
        if (!isJoined) {
            sql += `,\n[${t.name}] AS [${t.alias}]`;
        }
    });

    // NEW: Append WHERE
    const whereStr = buildLogicString(currentLogicState.WHERE);
    if(whereStr) {
        sql += `\nWHERE ${whereStr}`;
    }

    // NEW: Append HAVING
    const havingStr = buildLogicString(currentLogicState.HAVING);
    if(havingStr) {
        sql += `\nHAVING ${havingStr}`;
    }

    // NEW: Append ORDER BY
    if (orderBys.length > 0) {
        orderBys.sort((a,b) => a.ts - b.ts);
        sql += `\nORDER BY ${orderBys.map(o => o.val).join(", ")}`;
    }

    return sql;
}

function updateGeneratedSQL() {
    if(tablesOnCanvas.length === 0) {
        document.getElementById('sql-editor-area').value = "";
        return;
    }

    // 1. Identify CTEs on Canvas
    let directCTEs = tablesOnCanvas
        .filter(t => currentData.ctes[t.name])
        .map(t => t.name);
    
    // 2. Identify CTEs in Logic Deck (only ACTUAL CTEs needed for WITH clause)
    function traverseLogicForDeps(arr) {
        arr.forEach(item => {
            if(item.type === 'group') {
                traverseLogicForDeps(item.children);
            }
            else if(item.right && item.right.type === 'cte') {
                const tName = item.right.tableName || item.right.label.split('.')[0];
                if(currentData.ctes[tName]) directCTEs.push(tName);
                // We ignore regular tables here; they don't need a WITH clause definition
            }
        });
    }
    
    if(currentLogicState.WHERE) traverseLogicForDeps(currentLogicState.WHERE);
    if(currentLogicState.HAVING) traverseLogicForDeps(currentLogicState.HAVING);

    // 3. Resolve Dependencies and Order
    directCTEs = [...new Set(directCTEs)];
    let orderedCTEs = getResolvedCTEs(directCTEs);

    // 4. Generate WITH Header
    let cteHeader = "";
    if (orderedCTEs.length > 0) {
        let defs = orderedCTEs.map(name => {
            return `[${name}] AS (\n${currentData.ctes[name].sql}\n)`;
        });
        cteHeader = "WITH " + defs.join(",\n") + "\n";
    }

    // 5. Generate Main Query
    let mainSQL = generateSQLFromCanvas();
    document.getElementById('sql-editor-area').value = cteHeader + mainSQL;
}

// Topological Sort for CTEs
function getResolvedCTEs(neededNames) {
    let visited = new Set();
    let ordered = [];
    
    function visit(name) {
        if (visited.has(name)) return;
        visited.add(name);
        
        const cte = currentData.ctes[name];
        if (cte && cte.dependencies) {
            cte.dependencies.forEach(depName => visit(depName));
        }
        ordered.push(name);
    }

    neededNames.forEach(name => visit(name));
    return ordered;
}

function handleApply() {
    try {
        const sql = document.getElementById('sql-editor-area').value;
        let allBaseTables = new Set();
        
        tablesOnCanvas.forEach(t => {
            if (!currentData.ctes[t.name]) allBaseTables.add(t.name);
        });

        let directCTEs = tablesOnCanvas.filter(t => currentData.ctes[t.name]).map(t => t.name);
        
        function traverseLogicForDeps(arr) {
            if (!arr || !Array.isArray(arr)) return;
            arr.forEach(item => {
                if(item.type === 'group') traverseLogicForDeps(item.children);
                else if(item.right && item.right.type === 'cte') {
                    const tName = item.right.tableName || (item.right.label ? item.right.label.split('.')[0] : null);
                    if (tName) {
                        if(currentData.ctes[tName]) directCTEs.push(tName);
                        else allBaseTables.add(tName); // FIX: Capture Regular Table dependency in subquery
                    }
                }
            });
        }
        if (currentLogicState.WHERE) traverseLogicForDeps(currentLogicState.WHERE);
        if (currentLogicState.HAVING) traverseLogicForDeps(currentLogicState.HAVING);

        directCTEs = [...new Set(directCTEs)];
        let allNeededCTEs = getResolvedCTEs(directCTEs);

        allNeededCTEs.forEach(cteName => {
            const cte = currentData.ctes[cteName];
            if (cte && cte.baseTables) cte.baseTables.forEach(t => allBaseTables.add(t));
        });

        const finalTableList = [...allBaseTables];
        Office.context.ui.messageParent(JSON.stringify({ action: "runQuery", sql: sql, tables: finalTableList }));
    } catch (e) {
        console.error("Handle Apply Error:", e);
    }
}

// --- RENDER SIDEBAR ---
function renderIngredients() {
    const container = document.getElementById('ingredients-list');
    container.innerHTML = "";
    
    Object.keys(currentData.tables).forEach(tableName => {
        const isCTE = currentData.ctes[tableName] !== undefined;
        
        const group = document.createElement('div');
        group.className = 'table-group';
        
        const header = document.createElement('div');
        header.className = 'table-header';
        
        if (isCTE) {
            header.innerHTML = `<span style="color:#d13438">⚡ ${tableName}</span>`;
            header.style.borderLeft = "3px solid #d13438";
            
            // Actions container
            const actionsDiv = document.createElement('div');
            actionsDiv.style.marginLeft = "auto";
            actionsDiv.style.display = "flex";
            actionsDiv.style.gap = "8px";

            // EDIT BUTTON
            const editBtn = document.createElement('span');
            editBtn.innerText = "✎";
            editBtn.title = "Edit CTE";
            editBtn.style.cursor = "pointer";
            editBtn.style.fontSize = "12px";
            
            // UPDATED: Use showConfirm instead of window.confirm
            editBtn.onclick = (e) => {
                e.stopPropagation();
                showConfirm(
                    "Edit CTE",
                    `Load CTE '${tableName}' for editing? This will clear current canvas.`,
                    () => loadCTEForEdit(tableName)
                );
            };

            // DELETE BUTTON (NEW)
            const delBtn = document.createElement('span');
            delBtn.innerText = "✖";
            delBtn.title = "Delete CTE";
            delBtn.className = "delete-icon";
            delBtn.onclick = (e) => {
                e.stopPropagation();
                showConfirm(
                    "Delete CTE",
                    `Are you sure you want to delete '${tableName}'?`,
                    () => deleteCTE(tableName)
                );
            };

            actionsDiv.appendChild(editBtn);
            actionsDiv.appendChild(delBtn);
            header.appendChild(actionsDiv);

        } else {
            header.innerHTML = `<span>▼ ${tableName}</span>`;
        }
        
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

// ... (Rest of existing logic: createChip, toggleType, handleCanvasDrop, addTableToCanvas, renderTableCard, removeTable, join logic, etc. REMAIN UNCHANGED) ...
// Ensure you keep the existing helper functions below this point from the previous file state
// For brevity, I am not re-pasting functions that didn't change, but in a real file write, they exist here.

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
    
    const iconDiv = chip.querySelector('.chip-icon');
    if (isVirtual || type === "CALC") {
        iconDiv.onclick = (e) => { e.stopPropagation(); openModal(colName); };
        iconDiv.title = "Edit Formula";
    } else {
        iconDiv.onclick = (e) => { e.stopPropagation(); toggleType(colName, tableName); };
        iconDiv.title = "Toggle Type";
    }
    return chip;
}

function toggleType(colName, tableName) {
    const types = ["TEXT", "NUMBER", "DATE"];
    const targetTable = tableName || "Current Table";
    if (currentData.tables[targetTable]) {
        let current = currentData.tables[targetTable].colTypes[colName] || "TEXT";
        let idx = types.indexOf(current);
        let next = types[(idx + 1) % types.length];
        currentData.tables[targetTable].colTypes[colName] = next;
        renderIngredients();
    }
}

function handleCanvasDrop(e) {
    e.preventDefault();
    const rawData = e.dataTransfer.getData("application/json");
    if (!rawData) return;
    const data = JSON.parse(rawData);
    
    if (data.type === "table") {
        addTableToCanvas(data.name, e.clientX, e.clientY);
    } else if (data.type === "move_card") {
        const card = tablesOnCanvas.find(t => t.id === data.id);
        if (card) {
            const canvasRect = document.getElementById('visual-drop-zone').getBoundingClientRect();
            card.left = e.clientX - canvasRect.left - dragOffsetX;
            card.top = e.clientY - canvasRect.top - dragOffsetY;
            if(card.left < 0) card.left = 0;
            if(card.top < 0) card.top = 0;
            const cardEl = document.getElementById(card.id);
            cardEl.style.left = card.left + "px";
            cardEl.style.top = card.top + "px";
            renderJoins();
        }
    }
}

function addTableToCanvas(tableName, x, y) {
    if (tablesOnCanvas.find(t => t.name === tableName)) return;
    const canvasRect = document.getElementById('visual-drop-zone').getBoundingClientRect();
    let left = x ? (x - canvasRect.left) : (tablesOnCanvas.length * 220 + 20);
    let top = y ? (y - canvasRect.top) : 50;
    
    const tableObj = {
        id: "tbl_" + new Date().getTime(),
        name: tableName,
        alias: tableName, 
        columns: currentData.tables[tableName].schema,
        columnAliases: {}, 
        sorts: {}, // NEW: Store sorts { "ColName": { dir: "ASC", ts: 12345 } }
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
    
    const isCTE = currentData.ctes[tableObj.name] !== undefined;

    const header = document.createElement('div');
    header.className = 'card-header';
    if(isCTE) header.style.background = "#d13438"; 
    header.draggable = true;
    
    header.innerHTML = `
        <div style="display:flex; align-items:center; gap:5px; flex:1;">
            <span class="title-text">${tableObj.alias}</span>
            <span class="edit-alias" style="font-size:10px; cursor:pointer; opacity:0.7;">✎</span>
        </div>
        <span class="card-close" onclick="removeTable('${tableObj.id}')">✖</span>
    `;
    
    const titleSpan = header.querySelector('.title-text');
    const editBtn = header.querySelector('.edit-alias');
    
    editBtn.onclick = (e) => {
        e.stopPropagation(); 
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
        
        // Add sorting indicator class if sorted
        const sortInfo = tableObj.sorts && tableObj.sorts[col];
        if (sortInfo) item.classList.add('has-sort');
        if (tableObj.selected.includes(col)) item.classList.add('selected');

        const colAlias = tableObj.columnAliases && tableObj.columnAliases[col];
        
        // Container for text parts
        const textDiv = document.createElement('div');
        textDiv.style.flex = "1";
        textDiv.style.display = "flex";
        textDiv.style.flexDirection = "column"; // Default column for alias stacking
        
        if (colAlias) {
            textDiv.innerHTML = `
                <div style="line-height:1.2;">
                    <span class="col-original">${col}</span>
                    <span class="col-alias">${colAlias}</span>
                </div>`;
        } else {
            textDiv.innerText = col;
        }
        
        // Append Sort Badge if active
        if (sortInfo) {
            // Get Global Rank
            const rank = getGlobalSortRank(tableObj.id, col);
            const badge = document.createElement('span');
            badge.className = 'sort-badge';
            badge.innerText = `${sortInfo.dir} ${rank}`;
            // If aliased, append to alias line, otherwise straight after text
            if(colAlias) {
                // Find the bold alias part
                const aliasSpan = textDiv.querySelector('.col-alias');
                if(aliasSpan) aliasSpan.appendChild(badge);
            } else {
                textDiv.appendChild(badge);
                textDiv.style.flexDirection = "row"; // Row if single line
                textDiv.style.alignItems = "center";
            }
        }

        item.appendChild(textDiv);

        // --- NEW: Action Bar (Sort) ---
        const actionsDiv = document.createElement('div');
        actionsDiv.className = 'col-actions';
        
        const sortBtn = document.createElement('div');
        sortBtn.className = `btn-sort ${sortInfo ? 'active' : ''}`;
        sortBtn.innerText = "⇅";
        sortBtn.title = "Sort (ASC -> DESC -> OFF)";
        sortBtn.onclick = (e) => {
            e.stopPropagation();
            // Toggle Logic: None -> ASC -> DESC -> None
            if (!tableObj.sorts) tableObj.sorts = {};
            
            if (!sortInfo) {
                tableObj.sorts[col] = { dir: "ASC", ts: Date.now() };
            } else if (sortInfo.dir === "ASC") {
                tableObj.sorts[col].dir = "DESC";
            } else {
                delete tableObj.sorts[col];
            }
            
            // Re-render ALL cards to update global ranks
            // Clearing canvas DOM but keeping SVG structure is cleanest way without complex update logic
            const canvas = document.getElementById('visual-drop-zone');
            const svg = document.getElementById('connections-layer');
            canvas.innerHTML = '';
            canvas.appendChild(svg);
            tablesOnCanvas.forEach(t => renderTableCard(t));
            renderJoins();
            
            updateGeneratedSQL();
        };
        actionsDiv.appendChild(sortBtn);
        item.appendChild(actionsDiv);

        // Double click rename logic
        item.ondblclick = (e) => {
            e.stopPropagation();
            if(item.querySelector('input')) return;
            const currentDisplay = colAlias || col;
            
            // Hide display, show input
            textDiv.style.display = 'none';
            // Also hide actions during edit
            actionsDiv.style.display = 'none'; 
            
            item.draggable = false; 

            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentDisplay;
            input.className = 'alias-input';
            
            input.onclick = (ev) => ev.stopPropagation();
            input.onmousedown = (ev) => ev.stopPropagation();

            const finishEdit = () => {
                const newVal = input.value.trim();
                if (newVal && newVal !== col) {
                    if (!tableObj.columnAliases) tableObj.columnAliases = {};
                    tableObj.columnAliases[col] = newVal;
                } else {
                    if (tableObj.columnAliases && tableObj.columnAliases[col]) {
                        delete tableObj.columnAliases[col];
                    }
                }
                const currentCard = document.getElementById(tableObj.id);
                if(currentCard) currentCard.remove();
                renderTableCard(tableObj);
                updateGeneratedSQL();
            };

            input.onblur = finishEdit;
            input.onkeydown = (ev) => { if(ev.key === 'Enter') input.blur(); };
            item.insertBefore(input, actionsDiv); // Insert before actions
            input.focus();
        };

        item.dataset.table = tableObj.name; 
        item.dataset.col = col;
        item.draggable = true;

        item.onclick = (e) => {
            if(e.target.closest('.col-anchor')) return;
            if(e.target.tagName === 'INPUT') return;
            // FIX: Ignore clicks on action buttons to prevent selection toggle
            if(e.target.closest('.col-actions')) return; 

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
                type: "join_link", tableName: tableObj.name, alias: tableObj.alias, colName: col
            }));
        };

        item.ondragover = (e) => { e.preventDefault(); item.style.background = "#e6f2ff"; };
        item.ondragleave = (e) => { item.style.background = ""; };
        item.ondrop = (e) => handleJoinDrop(e, tableObj.name, col); 

        // FIX: Use appendChild instead of innerHTML+= to preserve event listeners on sort button
        const anchor = document.createElement('div');
        anchor.className = 'col-anchor';
        item.appendChild(anchor);

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

function handleJoinDrop(e, targetTable, targetCol) {
    e.preventDefault(); e.target.style.background = "";
    const raw = e.dataTransfer.getData("application/json");
    if (!raw) return;
    const src = JSON.parse(raw);
    if (src.type !== "join_link") return;
    if (src.tableName === targetTable) return; 

    const newJoin = {
        id: "j_" + new Date().getTime(),
        sourceTable: src.tableName, sourceCol: src.colName,
        targetTable: targetTable, targetCol: targetCol,
        type: "INNER JOIN", op: "="
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
            allCols.push({ name: col, type: currentData.tables[tableName].colTypes[col] || "TEXT" });
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

        chip.innerHTML = `<div class="chip-icon ${iconClass}" style="width:25px; font-size:10px;">${iconText}</div><div class="chip-name" style="padding:2px 6px;">${colData.name}</div>`;
        
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
        <span class="cond-del">✖</span>`;
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
    const tbl = tablesOnCanvas.find(t => t.id === id);
    if(tbl) {
        currentData.joins = currentData.joins.filter(j => j.sourceTable !== tbl.name && j.targetTable !== tbl.name);
    }
    tablesOnCanvas = tablesOnCanvas.filter(t => t.id !== id);
    document.getElementById(id).remove();
    renderJoins();
    updateGeneratedSQL();
};
