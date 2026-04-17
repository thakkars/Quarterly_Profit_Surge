<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Equity Research Terminal</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f4f7f9; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .header-area { display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; flex-wrap: wrap; gap: 15px; }
        h1 { color: #1a365d; margin: 0; }
        
        /* Action Button Styles */
        .action-panel { display: flex; align-items: center; gap: 10px; background: #edf2f7; padding: 10px 15px; border-radius: 8px; border: 1px solid #cbd5e0; flex-wrap: wrap;}
        #run-btn { background-color: #2b6cb0; color: white; border: none; padding: 10px 15px; border-radius: 6px; font-weight: bold; cursor: pointer; transition: 0.2s; font-size: 14px; }
        #run-btn:hover { background-color: #2c5282; }
        #run-btn:disabled { background-color: #a0aec0; cursor: not-allowed; }
        
        #export-btn { background-color: #4a5568; color: white; border: none; padding: 10px 15px; border-radius: 6px; font-weight: bold; cursor: pointer; transition: 0.2s; font-size: 14px; }
        #export-btn:hover { background-color: #2d3748; }

        #exceptions-btn { background-color: #e53e3e; color: white; border: none; padding: 10px 15px; border-radius: 6px; font-weight: bold; cursor: pointer; transition: 0.2s; font-size: 14px; }
        #exceptions-btn:hover { background-color: #c53030; }

        #workflow-status { font-size: 13px; font-weight: 600; color: #4a5568; margin-left: auto; }
        .spinner { display: inline-block; width: 14px; height: 14px; border: 2px solid rgba(0,0,0,0.1); border-radius: 50%; border-top-color: #2b6cb0; animation: spin 1s ease-in-out infinite; vertical-align: middle; margin-right: 5px; }
        @keyframes spin { to { transform: rotate(360deg); } }

        .controls { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 20px; padding: 15px; background: #edf2f7; border-radius: 8px; align-items: flex-start; }
        .filter-group { flex: 1; min-width: 250px; }
        .filter-group h3 { margin-top: 0; font-size: 16px; color: #2d3748; margin-bottom: 12px; border-bottom: 1px solid #cbd5e0; padding-bottom: 5px; }
        .checkbox-label { display: block; margin-bottom: 10px; cursor: pointer; font-size: 14px; font-weight: 500; }
        .checkbox-label input { margin-right: 8px; transform: scale(1.2); cursor: pointer;}
        
        /* Sleek Dropdown for Columns */
        .dropdown { width: 100%; }
        .dropdown-btn { background-color: #fff; border: 1px solid #cbd5e0; padding: 10px 15px; border-radius: 6px; font-size: 14px; font-weight: 600; color: #2b6cb0; cursor: pointer; width: 100%; text-align: left; display: flex; justify-content: space-between; align-items: center; transition: 0.2s; }
        .dropdown-btn:hover { background-color: #ebf8ff; border-color: #3182ce; }
        .dropdown-btn.active { border-radius: 6px 6px 0 0; border-bottom-color: transparent; }
        .dropdown-content { display: none; background-color: #fff; border: 1px solid #cbd5e0; border-top: none; border-radius: 0 0 6px 6px; padding: 15px; max-height: 250px; overflow-y: auto; }
        .dropdown-content.show { display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 10px; }
        .dropdown-content label { display: flex; align-items: center; cursor: pointer; font-size: 13px; color: #2d3748; }
        .dropdown-content input { margin-right: 6px; transform: scale(1.1); cursor: pointer; }
        
        /* Table Styles */
        .table-container { overflow-x: auto; border: 1px solid #e2e8f0; border-radius: 8px; max-height: 70vh;}
        table { width: 100%; border-collapse: collapse; font-size: 14px; background: white; }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #e2e8f0; white-space: nowrap; }
        th { background-color: #2b6cb0; color: white; position: sticky; top: 0; font-weight: 600; z-index: 10; cursor: pointer; transition: background 0.2s; user-select: none; }
        th:hover { background-color: #1a365d; }
        .sort-icon { font-size: 12px; margin-left: 5px; opacity: 0.7; }
        tr:hover { background-color: #ebf8ff; }
        .badge-yes { background-color: #48bb78; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
        .loading { text-align: center; padding: 40px; font-size: 18px; color: #718096; }
        .stats { text-align: right; font-size: 14px; color: #4a5568; margin-bottom: 10px; font-weight: 600; }
        .action-links { margin-top: 6px; display: flex; gap: 5px; }
        .action-links a { text-decoration: none; font-size: 11px; padding: 3px 6px; border-radius: 4px; background: #e2e8f0; color: #2b6cb0; font-weight: bold; transition: 0.2s; border: 1px solid #cbd5e0; }
        .action-links a:hover { background: #cbd5e0; color: #1a365d; border-color: #a0aec0;}

        /* EXCEPTIONS MODAL STYLES */
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; justify-content: center; align-items: center; backdrop-filter: blur(2px); }
        .modal-content { background: white; width: 90%; max-width: 800px; border-radius: 10px; box-shadow: 0 10px 25px rgba(0,0,0,0.2); display: flex; flex-direction: column; max-height: 85vh; overflow: hidden; }
        .modal-header { padding: 15px 20px; background: #1a365d; color: white; display: flex; justify-content: space-between; align-items: center; font-weight: bold; }
        .modal-close { background: none; border: none; color: white; font-size: 24px; cursor: pointer; }
        .modal-body { display: flex; flex: 1; overflow: hidden; }
        .modal-left { width: 40%; border-right: 1px solid #e2e8f0; overflow-y: auto; background: #f7fafc; padding: 10px; }
        .modal-right { width: 60%; overflow-y: auto; padding: 15px; background: white; }
        .group-item { padding: 10px; border-radius: 6px; cursor: pointer; display: flex; align-items: center; transition: 0.2s; margin-bottom: 5px; }
        .group-item:hover { background: #edf2f7; }
        .group-item.active { background: #e2e8f0; border-left: 4px solid #2b6cb0; font-weight: 600; }
        .group-item input { margin-right: 10px; transform: scale(1.2); cursor: pointer; }
        .sub-industry-item { padding: 8px 10px; display: flex; align-items: center; cursor: pointer; border-bottom: 1px solid #edf2f7; }
        .sub-industry-item:hover { background: #f7fafc; }
        .sub-industry-item input { margin-right: 10px; transform: scale(1.2); cursor: pointer; }
        .modal-footer { padding: 15px; background: #edf2f7; border-top: 1px solid #cbd5e0; text-align: right; }
        .btn-done { background: #2b6cb0; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: bold; }
        .btn-done:hover { background: #1a365d; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-area">
            <h1>📈 Equity Research Terminal</h1>
            
            <div class="action-panel">
                <button id="run-btn" onclick="triggerWorkflow()">▶️ Run Backend Scraper</button>
                <button id="exceptions-btn" onclick="openExceptionsModal()">🚫 Exceptions</button>
                <button id="export-btn" onclick="exportToCSV()">📥 Export CSV</button>
                <span id="workflow-status">Status: Idle</span>
            </div>
        </div>
        
        <div class="controls">
            <div class="filter-group">
                <h3>🎯 Strict Filters (Exclude Noise)</h3>
                <label class="checkbox-label"><input type="checkbox" id="filter-promoter"> Promoter Holding ≥ 50%</label>
                <label class="checkbox-label"><input type="checkbox" id="filter-pledged"> Pledged Shares ≤ 10%</label>
                <label class="checkbox-label"><input type="checkbox" id="filter-wc"> Working Capital Days ≤ 150</label>
                <label class="checkbox-label"><input type="checkbox" id="filter-growth"> 70% YoY Profit Surge (Latest Qtr)</label>
            </div>
            
            <div class="filter-group" style="flex: 2;">
                <h3>👁️ Column Settings</h3>
                <div class="dropdown">
                    <button onclick="toggleDropdown()" class="dropdown-btn" id="dropdown-btn-text">Select Columns ▼</button>
                    <div id="column-dropdown" class="dropdown-content">
                        </div>
                </div>
            </div>
        </div>

        <div class="stats" id="row-count">Loading data...</div>

        <div class="table-container">
            <table id="data-table">
                <thead id="table-head"></thead>
                <tbody id="table-body">
                    <tr><td class="loading" colspan="100%">Fetching live market data...</td></tr>
                </tbody>
            </table>
        </div>
    </div>

    <div id="exceptions-modal" class="modal-overlay">
        <div class="modal-content">
            <div class="modal-header">
                <span>🚫 Industry Exceptions (Check to Banish)</span>
                <button class="modal-close" onclick="closeExceptionsModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="modal-left" id="modal-groups">
                    </div>
                <div class="modal-right" id="modal-subs">
                    <div style="padding: 20px; color: #718096; text-align: center; margin-top: 50px;">
                        👈 Select an Industry Group on the left to view sub-industries.
                    </div>
                </div>
            </div>
            <div class="modal-footer">
                <span id="exception-count" style="margin-right: 15px; font-size: 14px; font-weight: bold; color: #e53e3e;">0 Industries Banished</span>
                <button class="btn-done" onclick="closeExceptionsModal()">Done</button>
            </div>
        </div>
    </div>

    <script>
        // --- CONFIGURATION ---
        const GITHUB_OWNER = "thakkars"; 
        const GITHUB_REPO = "Quarterly_Profit_Surge"; 
        const WORKFLOW_FILE = "run_screener.yml"; 
        const BRANCH = "main"; 
        const DATA_URL = "https://pub-9afc232687c44eebbb8cae359c7ac460.r2.dev/rawdata.json";
        
        // --- STATE VARIABLES ---
        let rawData = [];
        let currentFilteredData = []; 
        let allColumns = [];
        let visibleColumns = new Set();
        let matrixColumnName = "";
        let sortColumn = "Sr no"; 
        let sortAscending = true;

        // --- EXCEPTIONS STATE (THE BRAIN) ---
        // industryMap maps Group -> Array of Sub-Industries
        let industryMap = {}; 
        let activeModalGroup = null;
        
        // THIS IS THE MASTER SET. If a sub-industry string is in here, it is banished.
        let excludedIndustries = new Set(); 

        // --- HELPER TO RENAME COLUMNS ON SCREEN ---
        function getDisplayName(col) {
            if (col === "Company (Page)") return "Company";
            if (col === "date") return "Date";
            return col;
        }

        // ═══════════════════════════════════════════════════════════
        // EXCEPTIONS LOGIC & MODAL
        // ═══════════════════════════════════════════════════════════
        function buildIndustryMap() {
            industryMap = {};
            rawData.forEach(row => {
                let group = row["Industry Group"] || "Uncategorized";
                let sub = row["Industry"] || "Uncategorized";
                if (group.trim() === "" || group === "0") group = "Not Found";
                if (sub.trim() === "" || sub === "0") sub = "Not Found";
                
                if (!industryMap[group]) industryMap[group] = new Set();
                industryMap[group].add(sub);
            });
            // Convert Sets to sorted Arrays for the UI
            Object.keys(industryMap).forEach(group => {
                industryMap[group] = Array.from(industryMap[group]).sort();
            });
        }

        function openExceptionsModal() {
            document.getElementById("exceptions-modal").style.display = "flex";
            renderModalGroups();
        }

        function closeExceptionsModal() {
            document.getElementById("exceptions-modal").style.display = "none";
            applyFiltersAndRender(); // Re-render table when closed
        }

        function renderModalGroups() {
            const groupContainer = document.getElementById("modal-groups");
            groupContainer.innerHTML = "";
            
            let sortedGroups = Object.keys(industryMap).sort();

            sortedGroups.forEach(group => {
                const subs = industryMap[group];
                if(subs.length === 0) return;

                // Determine Checkbox State (Smart Selection)
                let excludedCount = subs.filter(sub => excludedIndustries.has(sub)).length;
                let isAllExcluded = excludedCount === subs.length;
                let isIndeterminate = excludedCount > 0 && excludedCount < subs.length;

                const div = document.createElement("div");
                div.className = `group-item ${activeModalGroup === group ? 'active' : ''}`;
                
                // Clicking the div opens the right panel. Clicking the checkbox toggles everything.
                div.innerHTML = `
                    <input type="checkbox" id="chk-grp-${group}" 
                        ${isAllExcluded ? 'checked' : ''} 
                        title="Ban entire group">
                    <div style="flex:1" onclick="selectModalGroup('${group}')">${group} <span style="font-size:11px; color:#718096">(${subs.length})</span></div>
                `;
                
                groupContainer.appendChild(div);

                // Set indeterminate visual state directly on element
                const cb = document.getElementById(`chk-grp-${group}`);
                if (isIndeterminate) cb.indeterminate = true;

                // Event listener for the group checkbox
                cb.addEventListener('change', (e) => {
                    const isChecked = e.target.checked;
                    subs.forEach(sub => {
                        if (isChecked) excludedIndustries.add(sub);
                        else excludedIndustries.delete(sub);
                    });
                    renderModalGroups(); // Re-render left
                    if (activeModalGroup === group) renderModalSubs(group); // Re-render right
                    updateExceptionCount();
                });
            });
            updateExceptionCount();
        }

        function selectModalGroup(group) {
            activeModalGroup = group;
            renderModalGroups(); // Re-render left to highlight active
            renderModalSubs(group);
        }

        function renderModalSubs(group) {
            const subContainer = document.getElementById("modal-subs");
            subContainer.innerHTML = `<h3 style="margin-top:0; color:#1a365d; border-bottom:2px solid #e2e8f0; padding-bottom:10px;">${group}</h3>`;
            
            const subs = industryMap[group];
            subs.forEach(sub => {
                const isExcluded = excludedIndustries.has(sub);
                
                const label = document.createElement("label");
                label.className = "sub-industry-item";
                label.innerHTML = `
                    <input type="checkbox" ${isExcluded ? 'checked' : ''} onchange="toggleExceptionSub('${sub}', this.checked, '${group}')">
                    <span>${sub}</span>
                `;
                subContainer.appendChild(label);
            });
        }

        function toggleExceptionSub(sub, isChecked, parentGroup) {
            if (isChecked) excludedIndustries.add(sub);
            else excludedIndustries.delete(sub);
            
            // Re-render the left panel to update the parent checkbox (Smart Selection)
            renderModalGroups();
            updateExceptionCount();
        }

        function updateExceptionCount() {
            document.getElementById("exception-count").innerText = `${excludedIndustries.size} Sub-Industries Banished`;
        }

        // ═══════════════════════════════════════════════════════════
        // UI INTERACTION (DROPDOWN FOR COLUMNS)
        // ═══════════════════════════════════════════════════════════
        function toggleDropdown() {
            const dropdown = document.getElementById("column-dropdown");
            const btn = document.getElementById("dropdown-btn-text");
            dropdown.classList.toggle("show");
            btn.classList.toggle("active");
        }

        function updateDropdownText() {
            document.getElementById("dropdown-btn-text").innerHTML = `👁️ ${visibleColumns.size} Columns Selected <span>▼</span>`;
        }

        window.onclick = function(event) {
            if (!event.target.matches('.dropdown-btn') && !event.target.closest('.dropdown-content') && !event.target.closest('.modal-overlay')) {
                const dropdown = document.getElementById("column-dropdown");
                const btn = document.getElementById("dropdown-btn-text");
                if (dropdown && dropdown.classList.contains('show')) {
                    dropdown.classList.remove('show');
                    btn.classList.remove('active');
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        // WORKFLOW TRIGGER LOGIC
        // ═══════════════════════════════════════════════════════════
        async function triggerWorkflow() {
            const btn = document.getElementById('run-btn');
            const statusText = document.getElementById('workflow-status');
            
            let token = localStorage.getItem('gh_token');
            if (!token) {
                token = prompt("Paste your GitHub Personal Access Token (PAT):");
                if (!token) return; 
                localStorage.setItem('gh_token', token.trim());
            }

            try {
                btn.disabled = true;
                statusText.innerHTML = `<span class="spinner"></span> Starting engine...`;

                const dispatchRes = await fetch(`https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`, {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Accept': 'application/vnd.github.v3+json',
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ ref: BRANCH })
                });

                if (!dispatchRes.ok) throw new Error("API Error");
                
                statusText.innerHTML = `<span class="spinner"></span> Engine spinning up...`;
                setTimeout(() => pollWorkflowStatus(token), 5000);

            } catch (error) {
                statusText.innerHTML = `❌ Error starting`;
                btn.disabled = false;
            }
        }

        async function pollWorkflowStatus(token) {
            const statusText = document.getElementById('workflow-status');
            try {
                const res = await fetch(`https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE}/runs?per_page=1`, {
                    headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/vnd.github.v3+json' }
                });
                const data = await res.json();
                
                if (data.workflow_runs && data.workflow_runs.length > 0) {
                    const latestRun = data.workflow_runs[0];
                    if (latestRun.status === "completed") {
                        if (latestRun.conclusion === "success") {
                            statusText.innerHTML = `✅ Complete! Fetching fresh data...`;
                            setTimeout(() => {
                                fetchData(); 
                                document.getElementById('run-btn').disabled = false;
                                statusText.innerHTML = `Status: Idle (Fresh Data Loaded)`;
                            }, 3000); 
                        } else {
                            statusText.innerHTML = `❌ Failed.`;
                            document.getElementById('run-btn').disabled = false;
                        }
                        return; 
                    } else {
                        statusText.innerHTML = `<span class="spinner"></span> Engine running...`;
                    }
                }
            } catch (error) { }
            setTimeout(() => pollWorkflowStatus(token), 10000);
        }

        // ═══════════════════════════════════════════════════════════
        // DATA FETCHING & CSV EXPORT
        // ═══════════════════════════════════════════════════════════
        async function fetchData() {
            try {
                const response = await fetch(`${DATA_URL}?t=${new Date().getTime()}`);
                if (!response.ok) throw new Error("Data not found");
                
                rawData = await response.json();
                
                if(rawData.length > 0) {
                    allColumns = Object.keys(rawData[0]).filter(col => col !== "Name");
                    matrixColumnName = allColumns.find(col => col.includes("(>70%)")) || "";

                    const defaultCols = ["Sr no", "date", "Company (Page)", "Industry", "Promoter Holding", "Pledged %", "WC Days"];
                    if (matrixColumnName) defaultCols.push(matrixColumnName);
                    
                    visibleColumns = new Set(defaultCols.filter(c => allColumns.includes(c)));
                    
                    buildIndustryMap(); // Build the Exceptions map dynamically!
                    renderColumnToggles();
                    applyFiltersAndRender();
                }
            } catch (error) {
                document.getElementById('table-body').innerHTML = `<tr><td class="loading" style="color:#e53e3e;" colspan="100%">Raw data not found. Run the backend scraper!</td></tr>`;
            }
        }

        function exportToCSV() {
            if (!currentFilteredData || currentFilteredData.length === 0) {
                alert("No data available to export.");
                return;
            }
            const orderedVisibleCols = allColumns.filter(col => visibleColumns.has(col));
            let csvLines = [];
            csvLines.push(orderedVisibleCols.map(c => `"${getDisplayName(c)}"`).join(','));
            
            currentFilteredData.forEach((row, index) => {
                let rowArray = orderedVisibleCols.map(col => {
                    let val = row[col];
                    if (col === "Sr no") val = index + 1;
                    if (val === null || val === undefined) val = "";
                    let strVal = String(val).replace(/"/g, '""'); 
                    return `"${strVal}"`;
                });
                csvLines.push(rowArray.join(','));
            });
            
            const csvData = new Blob([csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const csvUrl = URL.createObjectURL(csvData);
            const link = document.createElement('a');
            link.href = csvUrl;
            link.download = `Terminal_Export_${new Date().toISOString().slice(0,10)}.csv`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        // ═══════════════════════════════════════════════════════════
        // RENDERING (TABLE)
        // ═══════════════════════════════════════════════════════════
        function renderColumnToggles() {
            const container = document.getElementById('column-dropdown');
            container.innerHTML = '';
            
            allColumns.forEach(col => {
                const label = document.createElement('label');
                const isChecked = visibleColumns.has(col) ? 'checked' : '';
                label.innerHTML = `<input type="checkbox" value="${col}" ${isChecked} onchange="toggleColumn(this)"> <span>${getDisplayName(col)}</span>`;
                container.appendChild(label);
            });
            updateDropdownText();
        }

        function toggleColumn(checkbox) {
            if (checkbox.checked) visibleColumns.add(checkbox.value);
            else visibleColumns.delete(checkbox.value);
            updateDropdownText();
            applyFiltersAndRender();
        }

        function handleSort(column) {
            if (sortColumn === column) sortAscending = !sortAscending; 
            else { sortColumn = column; sortAscending = true; }
            applyFiltersAndRender();
        }

        function applyFiltersAndRender() {
            const filterPromoter = document.getElementById('filter-promoter').checked;
            const filterPledged = document.getElementById('filter-pledged').checked;
            const filterWc = document.getElementById('filter-wc').checked;
            const filterGrowth = document.getElementById('filter-growth').checked;

            let filteredData = rawData.filter(row => {
                // 1. Check Strict Checkbox Filters
                if (filterPromoter && (parseFloat(row["Promoter Holding"]) || 0) < 50) return false;
                if (filterPledged && (parseFloat(row["Pledged %"]) || 0) > 10) return false;
                if (filterWc && (parseFloat(row["WC Days"]) || 999) > 150) return false;
                if (filterGrowth && matrixColumnName && row[matrixColumnName] !== "Yes") return false;
                
                // 2. Check Exceptions Engine (Is this sub-industry banished?)
                let subIndustry = row["Industry"] || "Not Found";
                if (excludedIndustries.has(subIndustry)) return false;

                return true; 
            });

            if (sortColumn && sortColumn !== "Sr no") {
                filteredData.sort((a, b) => {
                    let valA = a[sortColumn]; let valB = b[sortColumn];
                    if (valA === null || valA === "") valA = -999999;
                    if (valB === null || valB === "") valB = -999999;
                    let numA = parseFloat(valA); let numB = parseFloat(valB);
                    if (!isNaN(numA) && !isNaN(numB)) return sortAscending ? numA - numB : numB - numA;
                    let strA = String(valA).toLowerCase(); let strB = String(valB).toLowerCase();
                    if (strA < strB) return sortAscending ? -1 : 1;
                    if (strA > strB) return sortAscending ? 1 : -1;
                    return 0;
                });
            } else if (sortColumn === "Sr no" && !sortAscending) {
                filteredData.reverse();
            }

            currentFilteredData = filteredData; 
            renderTable(filteredData);
        }

        function renderTable(data) {
            const thead = document.getElementById('table-head');
            const tbody = document.getElementById('table-body');
            document.getElementById('row-count').innerText = `Showing ${data.length} of ${rawData.length} stocks`;

            const orderedVisibleCols = allColumns.filter(col => visibleColumns.has(col));

            thead.innerHTML = '<tr>' + orderedVisibleCols.map(col => {
                let sortIcon = '';
                if (sortColumn === col) sortIcon = sortAscending ? '<span class="sort-icon">🔼</span>' : '<span class="sort-icon">🔽</span>';
                return `<th onclick="handleSort('${col}')" title="Click to sort">${getDisplayName(col)}${sortIcon}</th>`;
            }).join('') + '</tr>';

            if (data.length === 0) {
                tbody.innerHTML = `<tr><td colspan="${orderedVisibleCols.length}" class="loading">No stocks match the criteria.</td></tr>`;
                return;
            }

            tbody.innerHTML = data.map((row, index) => {
                return '<tr>' + orderedVisibleCols.map(col => {
                    if (col === "Sr no") {
                        let displayIndex = sortAscending ? (index + 1) : (data.length - index);
                        return `<td>${displayIndex}</td>`;
                    }
                    if (col === "Company (Page)") {
                        let fullName = row["Company (Page)"];
                        let shortName = encodeURIComponent(row["Name"]); 
                        return `<td>
                            <div style="font-weight: 700; color: #2d3748;">${fullName}</div>
                            <div class="action-links">
                                <a href="https://www.screener.in/company/${shortName}/" target="_blank" title="Open in Screener">🏢 Screener</a>
                                <a href="https://in.tradingview.com/chart/?symbol=NSE%3A${shortName}" target="_blank" title="Open in TradingView">📈 TV</a>
                            </div>
                        </td>`;
                    }
                    let val = row[col];
                    if (val === "Yes") return `<td><span class="badge-yes">Yes</span></td>`;
                    if (val === null || val === 0 || val === "") val = "-";
                    return `<td>${val}</td>`;
                }).join('') + '</tr>';
            }).join('');
        }

        document.querySelectorAll('.filter-group input[type="checkbox"]').forEach(cb => {
            cb.addEventListener('change', applyFiltersAndRender);
        });

        fetchData();
    </script>
</body>
</html>
