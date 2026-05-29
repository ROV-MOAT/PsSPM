<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE

.DESCRIPTION
    HTML Body and Mail
#>

#region HTML Template
$ExBottom = @"
<div class="fixed-bottom">
    <span id="summaryText" style="color: white; font-size: 14px;">Loading...</span>
    <span style="color: white; font-size: 12px; margin-left: auto;">$Version</span>
</div>
"@

$HtmlBodyR = $DataHtmlReport -replace '&lt;', '<' -replace '&#39;', "'" -replace '&gt;', '>'
$HtmlBody = $HtmlBodyR -join "`n"

$Global:FinalHtml = @"
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8" />
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/7.0.1/css/all.min.css"/>
<title>PsSPM Report</title>

<style>
    body {
        font-family: 'Trebuchet MS';
        margin: 0;
        padding: 0;
    }

    .fixed-header {
        display: flex;
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        background: #6d8196;
        border-radius: 0;
        z-index: 20;
        padding: 10px;
        border: 2px solid #ddd;
        border-bottom: none;
        width: 100%;
        box-sizing: border-box;
        max-height: 300px;
        align-items: center;
    }

    .fixed-bottom {
        display: flex;
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background: #6d8196;
        border-radius: 0;
        z-index: 20;
        padding: 5px 10px;
        border: 2px solid #ddd;
        width: 100%;
        box-sizing: border-box;
        height: 30px;
    }

    #rowFilter {
        visibility: collapse;
    }

    #rowFilter.visible {
        visibility: visible;
    }

    #rowFilter th {
        padding: 4px;
        background-color: white;
        cursor: default;
    }

    #rowFilter input {
        width: 100%;
        padding: 4px;
        box-sizing: border-box;
        border: 1px solid #ddd;
        border-radius: 3px;
        font-size: 12px;
        text-align: center;
    }

    #rowFilter input:focus {
        outline: none;
    }

    table {
        border-collapse: separate;
        width: 100%;
        border-spacing: 0;
        margin-bottom: 30px;
    }

    th, td {
        border: 2px solid #ddd;
        border-top: none;
        border-left: none;
        text-align: center;
        font-size: 13px;
        padding: 4px;
        width: 1%;
        box-sizing: border-box;
    }

    td {
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    th {
        font-weight: normal;
        background-color: #6d8196;
        color: white;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    th:first-child,
    td:first-child {
        border-left: 2px solid #ddd;
    }

    thead tr:first-child th {
        border-top: 2px solid #ddd;
    }
    
    tbody tr:hover { background-color: #f0f0f0; }
        
    ul {
        list-style-position: inside;
        list-style-type: disclosure-closed;
        text-align: Left;
        padding: 0;
        margin: 0;
        max-width: 400px;
        white-space: normal;
        word-wrap: break-word;
    }

    .toner-high { color: green; font-weight: bold; transition: all 0.3s ease; }
    .toner-medium { color: orange; font-weight: bold; transition: all 0.3s ease; }
    .toner-low { color: red; font-weight: bold; transition: all 0.3s ease; }

    .online { color: green; }
    .offline { color: red; }
    .error { color: orange; }

    a.printer-link { text-decoration: none; color: #0066cc; }

    .container { border-radius: 3px; padding: 5px; margin: 0; transition: all 0.3s ease; }
    .container:hover { background-color: black; background-size: cover; }
    .container:hover .toner-high,
    .container:hover .toner-medium,
    .container:hover .toner-low {
        color: white;
    }

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
    @keyframes fadeOut {
        0% { opacity: 1; }
        100% { opacity: 0; display: none; }
    }
    
    #btnExportExcel {
        width: fit-content;
        padding: 0 ;
        margin-left: 10px;
        background: #6d8196;
        color: white;
        border: none;
        cursor: pointer;
        text-align: left;
        font-size: 14px;
        transition: color 0.3s ease;
    }

    #btnExportExcel:hover {
        color: #FFDE21;
    }

    #btnToggleFilter {
        width: fit-content;
        padding: 0;
        margin-left: auto;
        background-color: #6d8196;
        color: white;
        border: none;
        cursor: pointer;
        text-align: left;
        font-size: 14px;
        transition: color 0.3s ease;
    }
    
    #scrollToTopBtn {
        position: fixed;
        bottom: 40px;
        right: 10px;
        width: 30px;
        height: 30px;
        background-color: #6d8196;
        color: white;
        border: none;
        border-radius: 3px;
        cursor: pointer;
        font-size: 14px;
        z-index: 1000;
        opacity: 0;
        padding: 0;
        margin: 0;
        visibility: hidden;
        transition: all 0.3s ease-in-out;
    }
        
    #scrollToTopBtn.show {
        opacity: 1;
        visibility: visible;

    }
        
    #scrollToTopBtn:hover {
        color: #FFDE21;
        transform: translateY(-5px);
    }
</style>
</head>
<body>
<button id="scrollToTopBtn" title="Up">
    <i class="fa-solid fa-angles-up fa-lg"></i>
</button>
<div class="fixed-header">
    <i class="fa-solid fa-print" style="color: white; margin-right: 5px;"></i>
    <span style="color: white; font-size: 14px;">This page was automatically generated $(Get-Date) by PowerShell SNMP printer monitoring (<a style='text-decoration: none; color: white;' href='https://github.com/ROV-MOAT/PsSPM' target='_blank'>PsSPM</a>).</span>
    <button id="btnToggleFilter">
        <i class="fa-solid fa-magnifying-glass fa-lg" style="margin-right: 5px;"></i><span>Filter</span>
    </button>
    <button id="btnExportExcel">
        <i class="fa-regular fa-file-excel fa-lg" style="margin-right: 5px;"></i><span>Export</span>
    </button>
</div>
<div class="spacer-div" style="margin-top: 40px;"></div>
<table id="reportTable">
    <thead>
        <tr>
            <th data-column="ip"><span>IP</span></th>
            <th data-column="status"><span>Status</span></th>
            <th data-column="name"><span>Name</span></th>
            <th data-column="mac"><span>MAC</span></th>
            <th data-column="model"><span>Model</span></th>
            <th data-column="serial"><span>S/N</span></th>
            <th><i class="fa-regular fa-file-lines"></i><span>Black</span></th>
            <th><i class="fa-regular fa-file-lines"></i><span>Color</span></th>
            <th><i class="fa-regular fa-file-lines"></i><span>Total</span></th>
            <th><span style="color:#00FFFF">C</span> Toner</th>
            <th><span style="color:#FD3DB5">M</span> Toner</th>
            <th><span style="color:#FFDE21">Y</span> Toner</th>
            <th><span style="color:#000000">K</span> Toner</th>
            <th><span style="color:#00FFFF">C</span><span style="color:#FD3DB5">M</span><span style="color:#FFDE21">Y</span><span style="color:#000000">K</span> DrumKit</th>
            <th><span style="color:#FFDE21">Display</span></th>
            <th><span style="color:#FFDE21">Active Alerts</span></th>
            <th data-column="errors"><span style="color:#FFDE21">E</span></th>
        </tr>
        <tr id="rowFilter">
            <th><input type="text" data-filter="ip" value=""></th>
            <th></th>
            <th><input type="text" data-filter="name" value=""></th>
            <th><input type="text" data-filter="mac" value=""></th>
            <th><input type="text" data-filter="model" value=""></th>
            <th><input type="text" data-filter="serial" value=""></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th><input type="text" data-filter="errors" value=""></th>
        </tr>
    </thead>
    <tbody>
    $HtmlBody
    </tbody>
</table>
$ExBottom

<script>
    document.addEventListener('DOMContentLoaded', () => {
        // ============================== DOM Elements ==============================
        const table = document.getElementById('reportTable');
        const tbody = table.querySelector('tbody');
        const summary = document.getElementById('summaryText');
        const filters = document.querySelectorAll('[data-filter]');
        const rowFilter = document.getElementById('rowFilter');
        const btnFilter = document.getElementById('btnToggleFilter');
        const btnExport = document.getElementById('btnExportExcel');
        const scrollToTopBtn = document.getElementById('scrollToTopBtn');
        
        let followToast = null;

        // ============================== Helpers ==============================
        const normalize = (value) => (value || '').toString().toLowerCase();

        // ============================== Column Mapping ==============================
        const columnMap = (() => {
            const map = {};
            table.querySelectorAll('thead th').forEach((th, index) => {
                if (th.hasAttribute('data-column')) {
                    map[th.dataset.column] = index;
                }
            });
            return map;
        })();

        // ============================== Row Cache ==============================
        const rows = [...tbody.querySelectorAll('tr')];
        const rowCache = rows.map(row => {
            const cells = row.children;
            const cache = {};
            for (const [field, index] of Object.entries(columnMap)) {
                const cell = cells[index];
                cache[field] = cell ? normalize(cell.textContent) : '';
            }
            return { row, cache };
        });

        // ============================== Summary Update ==============================
        const updateSummary = () => {
            if (!summary) return;

            const statusHeader = document.querySelector('th[data-column="status"]');
            const columnIndex = statusHeader.cellIndex;
            const visibleCount = rowCache.reduce((count, { row }) => count + (row.hidden ? 0 : 1), 0);
            
            let countOn = 0;
            let countOff = 0;
            let countErr = 0;
            
            rowCache.forEach(({ row }) => {
                const text = row.cells[columnIndex]?.textContent.trim();
                if (text === 'Online') countOn++;
                else if (text === 'Offline') countOff++;
                else if (text === 'Error') countErr++;
            });

            summary.textContent = 'Rows: ' + visibleCount + ' of ' + rowCache.length + ' | Status: Online ' + countOn + ' / Offline ' + countOff + ' / Error ' + countErr;
        };

        // ============================== Chunked Filtering ==============================
        const runChunkedFilter = (activeFilters) => {
            const filterKeys = Object.keys(activeFilters);
            const chunkSize = 400;
            let currentIndex = 0;

            const processChunk = () => {
                const endIndex = Math.min(currentIndex + chunkSize, rowCache.length);

                for (let i = currentIndex; i < endIndex; i++) {
                    const { row, cache } = rowCache[i];
                    let shouldHide = false;

                    for (const key of filterKeys) {
                        if (!cache[key].includes(activeFilters[key])) {
                            shouldHide = true;
                            break;
                        }
                    }
                    row.hidden = shouldHide;
                }

                currentIndex = endIndex;

                if (currentIndex < rowCache.length) {
                    requestAnimationFrame(processChunk);
                } else {
                    updateSticky();
                    updateSummary();
                }
            };

            requestAnimationFrame(processChunk);
        };

        const filterChunked = () => {
            const active = {};
            filters.forEach(filter => {
                const value = normalize(filter.value);
                if (value) {
                    active[filter.dataset.filter] = value;
                }
            });

            const filterKeys = Object.keys(active);

            if (filterKeys.length === 0) {
                rowCache.forEach(({ row }) => row.hidden = false);
                updateSummary();
                updateSticky();
                return;
            }

            runChunkedFilter(active);
        };

        // ============================== Debounce ==============================
        const debounce = (func, delay = 350) => {
            let timeoutId;
            return (...args) => {
                clearTimeout(timeoutId);
                timeoutId = setTimeout(() => func(...args), delay);
            };
        };

        const applyFilter = debounce(filterChunked, 350);

        // ============================== Filter Event Listeners ==============================
        filters.forEach(filter => filter.addEventListener('input', applyFilter));

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                filters.forEach(filter => filter.value = '');
                filterChunked();
            }
        });

        // ============================== Sticky Header ==============================
        const updateSticky = () => {
            const fixedHeader = document.querySelector('.fixed-header');
            const thead = table.querySelector('thead');
            if (!fixedHeader || !thead) return;

            const headerRows = [...thead.rows];
            if (headerRows.length === 0) return;

            const fixedHeight = fixedHeader.offsetHeight;
            const firstRowHeight = headerRows[0].offsetHeight;
            const secondRowHeight = headerRows[1]?.offsetHeight || 0;

            const spacerDiv = document.querySelector('.spacer-div');
            if (spacerDiv) {
                spacerDiv.style.marginTop = fixedHeight + 'px';
            }

            const setStickyStyles = (headers, topPosition, zIndexValue, backgroundColor) => {
                headers.forEach(header => {
                    Object.assign(header.style, {
                        position: 'sticky',
                        top: topPosition,
                        zIndex: zIndexValue,
                        background: backgroundColor
                    });
                });
            };

            const firstRowHeaders = headerRows[0].querySelectorAll('th');
            setStickyStyles(firstRowHeaders, fixedHeight + 'px', '15', '#6d8196');

            if (headerRows[1]) {
                const secondRowHeaders = headerRows[1].querySelectorAll('th');
                const topPosition = (fixedHeight + firstRowHeight) + 'px';
                setStickyStyles(secondRowHeaders, topPosition, '10', 'white');
            }

            document.documentElement.style.scrollPaddingTop = (fixedHeight + firstRowHeight + secondRowHeight) + 'px';
        };

        // ============================== Toast Notifications ==============================
        const createToast = () => {
            const toast = document.createElement('div');
            Object.assign(toast.style, {
                position: 'fixed',
                background: '#6d8196',
                color: '#fff',
                padding: '8px',
                borderRadius: '3px',
                fontSize: '13px',
                zIndex: 9999,
                pointerEvents: 'none',
                border: '1px solid rgba(255,255,255,0.1)',
                whiteSpace: 'nowrap',
                transition: 'opacity 0.15s ease',
                opacity: '0'
            });
            document.body.appendChild(toast);
            return toast;
        };

        const positionToast = (event, toast) => {
            const toastWidth = toast.offsetWidth;
            const toastHeight = toast.offsetHeight;
            
            let leftPosition = event.clientX - toastWidth - 10;
            let topPosition = event.clientY + 10;
            
            if (leftPosition < 5) leftPosition = event.clientX + 15;
            if (topPosition < 5) topPosition = 5;
            if (topPosition + toastHeight > window.innerHeight - 5) {
                topPosition = window.innerHeight - toastHeight - 5;
            }
            
            toast.style.left = leftPosition + 'px';
            toast.style.top = topPosition + 'px';
        };

        const showToast = (event, message) => {
            if (!followToast) {
                followToast = createToast();
            }
            followToast.innerHTML = message;
            followToast.style.opacity = '1';
            positionToast(event, followToast);
        };

        const hideToast = () => {
            if (followToast) {
                followToast.style.opacity = '0';
            }
        };

        // ============================== Tooltips ==============================
        document.querySelectorAll('[data-message]').forEach(element => {
            element.style.cursor = 'help';
            
            element.addEventListener('mouseenter', (event) => showToast(event, element.dataset.message));
            element.addEventListener('mousemove', (event) => {
                if (followToast && followToast.style.opacity === '1') {
                    positionToast(event, followToast);
                }
            });
            element.addEventListener('mouseleave', hideToast);
        });

        // ============================== Copy to Clipboard ==============================
        const markCopyable = () => {
            const copyableColumns = {};
            table.querySelectorAll('th[data-column]').forEach((th, index) => {
                const column = th.dataset.column;
                if (column === 'mac' || column === 'serial') {
                    copyableColumns[column] = index;
                }
            });

            for (const { row } of rowCache) {
                for (const columnIndex of Object.values(copyableColumns)) {
                    const cell = row.children[columnIndex];
                    if (!cell) continue;

                    const cellText = cell.textContent.trim();
                    if (!cellText) continue;

                    cell.style.cursor = 'copy';
                    cell.title = 'Click to copy';
                    cell.addEventListener('click', () => {
                        navigator.clipboard.writeText(cellText).then(() => {
                            const originalText = cell.textContent;
                            cell.textContent = 'Copied!';
                            cell.style.backgroundColor = '#e8f5e9';
                            setTimeout(() => {
                                cell.textContent = originalText;
                                cell.style.backgroundColor = '';
                            }, 800);
                        });
                    });
                }
            }
        };

        // ============================== Filter Toggle Button ==============================
        btnFilter.addEventListener('click', () => {
            rowFilter.classList.toggle('visible');
            btnFilter.style.color = rowFilter.classList.contains('visible') ? '#FFDE21' : 'white';
            updateSticky();
        });

        // ============================== Export to Excel ==============================
        btnExport.addEventListener('click', () => {
            const columnsToSkip = [14, 15, 16];
            let htmlContent = '<html><head><meta charset="UTF-8"><style>td,th{mso-number-format:"\\@"; text-align:"center"; vertical-align:"middle";}</style></head><body><table border="1">';

            const allRows = table.querySelectorAll('tr');
            allRows.forEach((row, index) => {
                if (row?.id === 'rowFilter') return;
                if (row.hidden) return;

                htmlContent += '<tr>';
                const cells = row.querySelectorAll('td, th');
                cells.forEach((cell, cellIndex) => {
                    if (!columnsToSkip.includes(cellIndex)) {
                        const tagName = index === 0 ? 'th' : 'td';
                        htmlContent += '<' + tagName + '>' + cell.textContent + '</' + tagName + '>';
                    }
                });
                htmlContent += '</tr>';
            });

            htmlContent += '</table></body></html>';

            const blob = new Blob([htmlContent], { type: 'application/vnd.ms-excel' });
            const blobUrl = URL.createObjectURL(blob);
            const downloadLink = Object.assign(document.createElement('a'), {
                href: blobUrl,
                download: 'PsSPM_Report.xls'
            });
            document.body.appendChild(downloadLink);
            downloadLink.click();
            downloadLink.remove();
            URL.revokeObjectURL(blobUrl);
        });

        // ============================== Scroll to Top Button ==============================
        if (scrollToTopBtn) {
            window.addEventListener('scroll', () => {
                scrollToTopBtn.classList.toggle('show', window.scrollY > 300);
            });

            scrollToTopBtn.addEventListener('click', () => {
                window.scrollTo({
                    top: 0,
                    behavior: 'smooth'
                });
            });
        }

        // ============================== Font Awesome Check ==============================
        const hasFontAwesome = () => {
            const testIcon = document.createElement('i');
            testIcon.className = 'fas fa-home';
            testIcon.style.cssText = 'position: absolute; visibility: visible; display: inline-block;';
            document.body.appendChild(testIcon);
            
            const rect = testIcon.getBoundingClientRect();
            const isEmpty = rect.width === 0 && rect.height === 0;
            
            testIcon.remove();
            
            return !isEmpty;
        };
        
        // ============================== Icon Handler ==============================
        const initAllIcons = () => {
            if (!hasFontAwesome()) {
                document.querySelectorAll('#reportTable th').forEach(th => {
                    const icon = th.querySelector('i');
                    const span = th.querySelector('span');
                    
                    if (icon && span) {
                        const text = span.textContent;
                        icon.remove();
                        span.insertAdjacentHTML('beforebegin', '📄 ');
                    }
                });
                
                const replaceIcons = (mappings) => {
                    mappings.forEach(({ target, selector, emoji, margin = true }) => {
                        const element = typeof target === 'string' ? document.querySelector(target) : target;
                        const icon = element?.querySelector(selector);
                        
                        if (icon) {
                            const span = document.createElement('span');
                            span.textContent = emoji;
                            if (margin) span.style.marginRight = '5px';

                            icon.replaceWith(span);
                        }
                    });
                };

                replaceIcons([
                    { target: '.fixed-header', selector: '.fa-print', emoji: '🖨️' },
                    { target: btnFilter, selector: '.fa-magnifying-glass', emoji: '🔍' },
                    { target: btnExport, selector: '.fa-file-excel', emoji: '💾' },
                    { target: scrollToTopBtn, selector: '.fa-angles-up', emoji: '▲', margin: false }
                ]);
            }
        };

        // ============================== Initialization ==============================
        window.addEventListener('load', updateSticky);
        window.addEventListener('resize', updateSticky);

        initAllIcons();
        markCopyable();
        updateSummary();
    });
</script>
</body>
</html>
"@

$Global:MailHtmlBody = @"
<span><font color="black">This message was automatically generated by PowerShell SNMP printer monitoring (PsSPM).</font></span>
<p>Date: $(Get-Date)</p>
"@
#endregion