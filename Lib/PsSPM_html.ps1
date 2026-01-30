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
<style>
    body { font-family: 'Trebuchet MS', sans-serif; margin: 20px; }
    table {
        border-collapse: collapse;
        border: 1px solid black;
        font-size: 90%;
        width: 100%;
        margin-bottom: 20px;
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
    
    /* Search */
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
        transition: all 0.3s ease;
    }
    .search-input:focus {
        outline: none;
        border: 2px solid green;
    }

    .column-selector {
        font-size: 14px;
    }
    .column-option {
        display: inline-block;
        margin-left: 10px;
        padding: 5px 10px;
        background-color: #e7e7e7;
        border-radius: 3px;
        cursor: pointer;
        transition: background-color 0.3s;
    }
    .column-option:hover {
        background-color: #d4d4d4;
    }
    .column-option.active {
        background-color: green;
        color: white;
    }

    tr.hidden {
        display: none;
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
    
    // Basic filtering function
    function filterTable() {
        const searchTerm = document.getElementById('searchInput').value.trim();
        const table = document.getElementById('PrinterTable');
        const rows = table.getElementsByTagName('tr');
        
        // If the search query is empty, we show all lines
        if (!searchTerm) {
            for (let i = 1; i < rows.length; i++) {
                rows[i].classList.remove('hidden');
            }
            noResults.style.display = 'none';
            return;
        }
        
        const lowerSearchTerm = searchTerm.toLowerCase();
        let foundResults = false;
                
        // Filtering rows
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagName('td');
            let shouldShow = false;
            
            if (selectedColumn === 'all') {
                // Search in all columns
                for (let j = 0; j < cells.length; j++) {
                    const cellText = cells[j].textContent.toLowerCase();
                    if (cellText.includes(lowerSearchTerm)) {
                        shouldShow = true;
                    }
                }
            } else {
                // Search only in the selected column
                const columnIndex = parseInt(selectedColumn);
                if (cells[columnIndex]) {
                    const cellText = cells[columnIndex].textContent.toLowerCase();
                    if (cellText.includes(lowerSearchTerm)) {
                        shouldShow = true;
                    }
                }
            }
            
            if (shouldShow) {
                rows[i].classList.remove('hidden');
                foundResults = true;
            } else {
                rows[i].classList.add('hidden');
            }
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
    
    // Initialization
    document.addEventListener('DOMContentLoaded', function() {
        // We save the original data
        saveOriginalData();
        
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
        
    });
</script>
<span><font color="black">This page was automatically generated by PowerShell SNMP printer monitoring (<a style='text-decoration: none; color: #000000ff;' href='https://github.com/ROV-MOAT/PsSPM' target='_blank'>PsSPM</a>).</font></span>
<p></p>
<div class="search-container">
    <div class="column-selector">
        <strong>Search in:</strong>
        <div class="column-option active" data-column="all">All</div>
        <div class="column-option" data-column="0">IP</div>
        <div class="column-option" data-column="2">Name</div>
        <div class="column-option" data-column="3">MAC</div>
        <div class="column-option" data-column="4">Model</div>
        <div class="column-option" data-column="5">S/N</div>
        <input type="text" id="searchInput" placeholder="Enter text to search..." class="search-input">
    </div>
</div>
"@

$MailHtmlBody = @"
<span><font color="black">This message was automatically generated by PowerShell SNMP printer monitoring (PsSPM).</font></span>
<p>Date: $(Get-Date)</p>
"@

#endregion
