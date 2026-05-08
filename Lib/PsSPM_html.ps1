<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE

.DESCRIPTION
    HTML Body and Mail
#>

#region HTML Template
$Header = @"
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/7.0.1/css/all.min.css">
<style>
    body { font-family: 'Trebuchet MS', sans-serif; margin: 20px; }
    table {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 15px;
    }
    th, td {
        border: 2px solid #ddd;
        padding: 5px;
        text-align: center;
        font-weight: normal;
        font-size: 15px;
    }
    th {
        font-size: 17px;
        background-color: #6d8196;
        color: white;
        position: sticky;
        top: 0;
        cursor: pointer;
    }
    tr:hover { background-color: #f0f0f0; }
    li { margin-top: 5px; margin-bottom: 5px; }
    .online { color: green; }
    .offline { color: red; }
    .error { color: orange; }

    .toner-high { color: green; font-weight: bold; transition: all 0.3s ease; }
    .toner-medium { color: orange; font-weight: bold; transition: all 0.3s ease; }
    .toner-low { color: red; font-weight: bold; transition: all 0.3s ease; }

    a.printer-link { text-decoration: none; color: #0066cc; }
    a.show-link {
        display: inline-block;
        text-decoration: none;
        color: #0066cc;
        animation: show-link ease-in-out 1s infinite alternate;
    }
    @keyframes show-link {
        0% {
            transform: rotate(0deg);
        }
        25% {
            transform: rotate(5deg);
        }
        75% {
            transform: rotate(-5deg);
        }
        100% {
            transform: rotate(0deg);
        }
    }

    .tooltip { position: relative; display: inline-block; }

    /* Tooltip text */
    .tooltip .tooltiptext {
        list-style-position: inside;
        list-style-type: disclosure-closed;
        visibility: hidden;
        max-width: 400px;
        width: max-content; /* Allows the tooltip to size based on content up to max-width */
        white-space: normal; /* Ensures text wraps within the tooltip */
        word-wrap: break-word; /* Prevents long words from overflowing */
        background-color: #6d8196;
        color: #ffffff;
        text-align: Left;
        padding: 5px 5px 7px 5px;
        margin: 0;
        border-radius: 5px;
        opacity: 0;
        transition: visibility 0.3s ease, opacity 0.3s ease, background-color 0.3s ease;
 
        /* Position the tooltip text */
        position: absolute;
        z-index: 1000;
        right: 35px;
        top: 20px;
    }

    /* Show the tooltip text when you mouse over the tooltip container */
    .tooltip:hover .tooltiptext { visibility: visible; opacity: 1; }
    .tooltiptext:hover { background-color: #000000; }

    .container { border-radius: 3px; padding: 5px; margin: 0; text-align: center; transition: all 0.3s ease; }
    .container:hover { background-color: #000000; background-size: cover; }
    .container:hover .toner-high,
    .container:hover .toner-medium,
    .container:hover .toner-low {
        color: #ffffff;
    }

    .search-container {
        background-color: #ffffff;
        padding: 10px;
        border-radius: 5px;
        border: 2px solid #ddd;
        margin-bottom: 15px;
        width: max-content;
    }
    .search-input {
        padding: 4px;
        width: 220px;
        border: 2px solid #ddd;
        border-radius: 4px;
        font-size: 16px;
        margin-left: 10px;
        transition: border 0.3s ease;
    }
    .search-input:hover {
        border: 2px solid #6d8196;
    }
    .search-input:focus {
        outline: none;
        border: 2px solid #6d8196;
    }

    .column-selector {
        font-size: 14px;
    }
    .column-option {
        display: inline-block;
        margin-left: 10px;
        padding: 5px 10px;
        background-color: #f7f6f6;
        border-radius: 3px;
        color: black;
        cursor: pointer;
        transition: background-color 0.3s ease, color 0.3s ease;
        min-width: 40px;
        text-align: center;
    }
    .column-option:hover {
        color: white;
        background-color: #6d8196;
    }
    .column-option.active {
        color: white;
        background-color: #6d8196;
    }

    tr.hidden {
        display: none;
    }

    @keyframes fadeOut {
        0% { opacity: 1; }
        100% { opacity: 0; display: none; }
    }
    
    .exp-button {
        display: inline-block;
        text-decoration: none;
        min-width: 90px;
        font-family: inherit;
        appearance: none;
        border-radius: 3px;
        border: 0;
        background: #f7f6f6;
        color: black;
        font-size: 14px;
        cursor: pointer;
        padding: 5px 10px;
        transition: background-color 0.3s ease, color 0.3s ease;
    }
    .exp-button:hover {
        color: white;
        background-color: #6d8196;
    }
    .exp-button:focus {
        outline: none;
    }
</style>

<script>
    let selectedColumn = 'all';
    let originalTableData = [];
    
    // Save the original table data
    function saveOriginalData() {
        const table = document.getElementById('PrinterTable');
        const rows = table.getElementsByTagName('tr');
        originalTableData = [];
        
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagName('td');
            const rowData = [];
            for (let j = 0; j < cells.length; j++) {
                rowData.push(cells[j].innerHTML);
            }
            originalTableData.push(rowData);
        }
    }
    
    function addNoResultsMessage() {
        if (document.getElementById('noResults')) return;
        
        const table = document.getElementById('PrinterTable');
        if (!table) return;
        
        const noResultsDiv = document.createElement('div');
        noResultsDiv.id = 'noResults';
        noResultsDiv.style.cssText = 'display: none; text-align: center; padding: 20px; font-size: 18px; color: #6d8196; background: #f9f9f9; border-radius: 5px; margin-bottom: 15px;';
        noResultsDiv.innerHTML = '<i class="fa-solid fa-magnifying-glass"></i> Not found';

        table.parentNode.insertBefore(noResultsDiv, table.nextSibling);
    }

    function filterTable() {
        const searchTerm = document.getElementById('searchInput').value.trim();
        const table = document.getElementById('PrinterTable');
        if (!table) return;
        
        const rows = table.getElementsByTagName('tr');
        let noResults = document.getElementById('noResults');
        
        if (!noResults) {
            addNoResultsMessage();
            noResults = document.getElementById('noResults');
        }
        
        if (!searchTerm) {
            for (let i = 1; i < rows.length; i++) {
                if (rows[i]) rows[i].classList.remove('hidden');
            }
            if (noResults) noResults.style.display = 'none';
            return;
        }
        
        const lowerSearchTerm = searchTerm.toLowerCase();
        let visibleCount = 0;
        
        // Filter
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;
            
            const cells = row.getElementsByTagName('td');
            let shouldShow = false;
            
            if (selectedColumn === 'all') {
                for (let j = 0; j < cells.length; j++) {
                    if (cells[j].textContent.toLowerCase().includes(lowerSearchTerm)) {
                        shouldShow = true;
                        break;
                    }
                }
            } else {
                const columnIndex = parseInt(selectedColumn);
                if (cells[columnIndex] && cells[columnIndex].textContent.toLowerCase().includes(lowerSearchTerm)) {
                    shouldShow = true;
                }
            }
            
            if (shouldShow) {
                row.classList.remove('hidden');
                visibleCount++;
            } else {
                row.classList.add('hidden');
            }
        }
        
        if (noResults) {
            noResults.style.display = visibleCount === 0 ? 'block' : 'none';
        }
    }
    
    // Selecting a column to search
    function selectColumn(column) {
        selectedColumn = column;
        
        // Updating the visual state of buttons
        const options = document.querySelectorAll('.column-option');
        options.forEach(option => {
            if (option.dataset.column === column) {
                option.classList.add('active');
            } else {
                option.classList.remove('active');
            }
        });
        
        // Apply the filter with new settings
        filterTable();
    }

    function exportToExcelHTML(tableId, filename = 'export.xls', excludeColumns = []) {
        const table = document.getElementById(tableId);
        
        let html = '<html><head><meta charset="UTF-8">';

        html += '<style>td, th { mso-number-format: "\\@"; }</style>';
        html += '</head><body>';
        html += '<table border="1">';
        
        const allRows = table.querySelectorAll('tr');
        let isFirstVisibleRow = true;
        
        allRows.forEach(row => {
            const isVisible = row.offsetParent !== null && 
                            window.getComputedStyle(row).display !== 'none';
            
            if (isVisible) {
                html += '<tr style="text-align: left">';
                const cells = row.querySelectorAll('td, th');
                
                cells.forEach((cell, index) => {
                    if (!excludeColumns.includes(index)) {
                        const tag = (isFirstVisibleRow) ? 'th' : 'td';
                        html += '<' + tag + '>' + cell.innerText + '</' + tag + '>';
                    }
                });
                
                html += '</tr>';
                isFirstVisibleRow = false;
            }
        });
        
        html += '</table></body></html>';
        
        const blob = new Blob(['\uFEFF' + html], { type: 'application/vnd.ms-excel' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
    }
    
    //Copy text on click function

    function fallbackCopy(element) {
        const text = element.innerText || element.textContent;
        const textArea = document.createElement('textarea');
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        showNotification('Copied (fallback): ' + text);
    }

    async function copyText(element) {
        try {
            const text = element.innerText || element.textContent;
            await navigator.clipboard.writeText(text);
            showNotification('Сopied: ' + text);
        } catch (err) {
            console.error('Сopy error:', err);
            fallbackCopy(element);
        }
    }

    function showNotification(message) {
        const notification = document.createElement('div');
        notification.textContent = message;
        notification.style.cssText = ' position: fixed; bottom: 20px; right: 20px; background: #6d8196; color: white; font-size: 13px; padding: 10px 20px; border-radius: 5px; animation: fadeOut 2s forwards; ';
        document.body.appendChild(notification);
    
        setTimeout(() => notification.remove(), 2000);
    }

    // Initialization
    document.addEventListener('DOMContentLoaded', function() {
        // We save the original data
        saveOriginalData();
        addNoResultsMessage();
        
        // Handler for the search field
        const searchInput = document.getElementById('searchInput');
        searchInput.addEventListener('input', filterTable);
        searchInput.addEventListener('keyup', function(e) {
            if (e.key === 'Escape') {
                this.value = '';
                filterTable();
            }
        });
        
        // Handlers for selecting columns
        const columnOptions = document.querySelectorAll('.column-option');
        columnOptions.forEach(option => {
           option.addEventListener('click', function() {
                selectColumn(this.dataset.column);
            });
        });

        const headers = document.querySelectorAll('#PrinterTable th');

        const allIndexes = new Set();
        [1].forEach(i => allIndexes.add(i));
        for (let i = 6; i <= 16; i++) {
            allIndexes.add(i);
        }

        headers.forEach((header, index) => {
            header.addEventListener('click', function() {
                const columnValue = allIndexes.has(index) ? 'all' : index.toString();
                selectColumn(columnValue);
            });
        });
    });
</script>
"@

$ExBody = @"
<div style="border: 2px solid #ddd; border-radius: 5px; background-color: #6d8196; padding: 10px; margin-bottom: 15px;">
    <span><font color="white">This page was automatically generated $(Get-Date) by PowerShell SNMP printer monitoring (<a style='text-decoration: none; color: white;' href='https://github.com/ROV-MOAT/PsSPM' target='_blank'>PsSPM</a>).</font></span>
</div>
<div class="search-container">
    <div class="column-selector">
        <strong>Search in:</strong>
        <div class="column-option active" data-column="all">Table</div>
        <div class="column-option" data-column="0">IP</div>
        <div class="column-option" data-column="2">Name</div>
        <div class="column-option" data-column="3">MAC</div>
        <div class="column-option" data-column="4">Model</div>
        <div class="column-option" data-column="5">S/N</div>
        <input type="text" id="searchInput" placeholder="Enter text to search..." class="search-input">
    </div>
</div>
<div style="margin-bottom: 15px;">
    <button class="exp-button" onclick="exportToExcelHTML('PrinterTable', 'PsSPM.xls', [14, 15, 16])">
        <i class="fas fa-download me-1"></i> Export
    </button>
</div>
"@

$ExBottom = @"
<div><span style='font-size: 12px;'>$Version</span></div>
"@

$MailHtmlBody = @"
<span><font color="black">This message was automatically generated by PowerShell SNMP printer monitoring (PsSPM).</font></span>
<p>Date: $(Get-Date)</p>
"@
#endregion