/* global Office, document, console */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Init logic
        document.getElementById('btn-close').onclick = () => Office.context.ui.messageParent(JSON.stringify({ action: "close" }));
        
        // Listen for data from Parent (TaskPane)
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent);
    }
});

function onMessageFromParent(arg) {
    const message = JSON.parse(arg.message);
    if (message.type === "init") {
        renderIngredients(message.schema, message.virtuals);
    }
}

function renderIngredients(schema, virtuals) {
    const container = document.getElementById('ingredients-list');
    container.innerHTML = "";

    // 1. Current Table Group
    const group = document.createElement('div');
    group.className = 'table-group';
    group.innerHTML = `<div class="table-header">▼ Current Table</div>`;
    
    schema.forEach(col => {
        const chip = document.createElement('div');
        chip.className = 'chip';
        chip.draggable = true; // For later drag logic
        chip.innerHTML = `
            <div class="chip-icon" style="color:#0078d4">COL</div>
            <div class="chip-name">${col}</div>
        `;
        group.appendChild(chip);
    });
    container.appendChild(group);

    // 2. Virtuals
    if (Object.keys(virtuals).length > 0) {
        const vGroup = document.createElement('div');
        vGroup.className = 'table-group';
        vGroup.innerHTML = `<div class="table-header" style="color:#813a7c">▼ Calculated</div>`;
        Object.keys(virtuals).forEach(key => {
            const chip = document.createElement('div');
            chip.className = 'chip';
            chip.innerHTML = `
                <div class="chip-icon" style="color:#813a7c">ƒx</div>
                <div class="chip-name">${key}</div>
            `;
            vGroup.appendChild(chip);
        });
        container.appendChild(vGroup);
    }
}

// Global Tab Switcher
window.switchTab = function(tabName) {
    // UI toggle
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.getElementById('tab-' + tabName).classList.add('active');

    // View toggle
    document.getElementById('view-visual').style.display = 'none';
    document.getElementById('view-sql').style.display = 'none';
    document.getElementById('view-' + tabName).style.display = (tabName === 'sql' ? 'flex' : 'flex');
};