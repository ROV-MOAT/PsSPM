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
    <span style="color: white; font-size: 14px;" id="summaryText">Loading...</span>
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
        width: 100%;
        box-sizing: border-box;
        max-height: 300px;
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

    .btn-export {
        width: fit-content;
        padding: 0 ;
        background: #6d8196;
        color: white;
        border: none;
        cursor: pointer;
        text-align: left;
        font-size: 14px;
    }

    .btn-toggle-filters {
        width: fit-content;
        padding: 0;
        background-color: #6d8196;
        color: white;
        border: none;
        cursor: pointer;
        text-align: left;
        font-size: 14px;
    }

    .filter-row th {
        padding: 4px;
        background-color: transparent;
        cursor: default;
    }

    .filter-row input {
        width: 100%;
        padding: 4px;
        box-sizing: border-box;
        border: 1px solid #ddd;
        border-radius: 3px;
        font-size: 12px;
        text-align: center;
    }

    .filter-row input:focus {
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
        width: max-content;
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
        transition: all 0.3s ease;
        box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        z-index: 1000;
        /* Плавное появление */
        opacity: 0;
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
<button id="scrollToTopBtn" title="Up">↑</button>
<div class="fixed-header">
    <i style="color: white; margin-right: 5px;" class="fa-solid fa-print"></i>
    <span style="color: white; font-size: 14px;">This page was automatically generated $(Get-Date) by PowerShell SNMP printer monitoring (<a style='text-decoration: none; color: white;' href='https://github.com/ROV-MOAT/PsSPM' target='_blank'>PsSPM</a>).</span>
    <button style="margin-left: auto;"  id="btnToggleFilter" class="btn-toggle-filters">
        <i class="fa-solid fa-magnifying-glass fa-lg"></i> Filter
    </button>
    <button style="margin-left: 5px;" id="btnExportExcel" class="btn-export">
        <i class="fa-regular fa-file-excel fa-lg"></i> Export
    </button>
</div>
<div class="spacer-div" style="margin-top: 40px;"></div>
<table id="reportTable">
    <thead>
        <tr>
            <th data-column="ip"><span>IP</span></th>
            <th data-column="ping"><span>Ping</span></th>
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
        <tr id="rowFilter" class="filter-row" style="display: none;">
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

    const table      = document.getElementById('reportTable');
    const tbody      = table.querySelector('tbody');
    const summary    = document.getElementById('summaryText');
    const filters    = document.querySelectorAll('[data-filter]');
    const rowFilter  = document.getElementById('rowFilter');
    const btnFilter  = document.getElementById('btnToggleFilter');
    const btnExport  = document.getElementById('btnExportExcel');
    const rows       = [...tbody.querySelectorAll('tr')];
    let followToast  = null;

    const norm = v => (v || '').toString().toLowerCase();

    const colMap = (() => {
        const map = {};
        table.querySelectorAll('thead th')
            .forEach((th, i) => {
                if (th.hasAttribute('data-column')) {
                    map[th.dataset.column] = i;
                }
        });
        return map;
    })();

    const rowCache = rows.map(tr => {
        const cells = tr.children;
        const cache = {};
        for (const [field, idx] of Object.entries(colMap)) {
            const cell = cells[idx];
            cache[field] = cell ? norm(cell.textContent) : '';
        }
        return { tr, cache };
    });

    const updateSummary = () => {
        let visible = 0;
        for (const { tr } of rowCache) {
            if (!tr.hidden) visible++;
        }

        summary.textContent = "Printers: " + visible + " of " + rowCache.length;
    };

    function runChunkedFilter(activeFilters) {
        const keys = Object.keys(activeFilters);
        const size = 400;
        let index = 0;

        function processChunk() {
            const end = Math.min(index + size, rowCache.length);

            for (let i = index; i < end; i++) {
                const { tr, cache } = rowCache[i];
                let hide = false;

                for (const k of keys) {
                    if (!cache[k].includes(activeFilters[k])) {
                        hide = true;
                        break;
                    }
                }
                tr.hidden = hide;
            }
            index = end;

            if (index < rowCache.length) {
                requestAnimationFrame(processChunk);
            } else {
                updateSticky();
                updateSummary();
            }
        }
        requestAnimationFrame(processChunk);
    }

    function filterChunked() {
        const active = {};
        filters.forEach(f => {
            const v = norm(f.value);
            if (v) active[f.dataset.filter] = v;
        });

        const keys = Object.keys(active);

        if (!keys.length) {
            rowCache.forEach(({ tr }) => tr.hidden = false);
            updateSummary();
            return;
        }

        runChunkedFilter(active);
    }

    const debounce = (fn, d = 350) => {
        let t;
        return (...a) => {
            clearTimeout(t);
            t = setTimeout(() => fn(...a), d);
        };
    };

    const applyFilter = debounce(filterChunked, 350);

    filters.forEach(f => f.addEventListener('input', applyFilter));
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') {
            filters.forEach(f => f.value = '');
            filterChunked();
        }
    });

    const updateSticky = () => {
        const fixed = document.querySelector('.fixed-header');
        const thead = table.querySelector('thead');
        if (!fixed || !thead) return;

        const headRows = [...thead.rows];
        if (!headRows.length) return;

        const divH    = fixed.offsetHeight;
        const firstH  = headRows[0].offsetHeight;
        const secondH = headRows[1]?.offsetHeight || 0;

        document.querySelector('.spacer-div').style.marginTop = divH + 'px';

        headRows[0].querySelectorAll('th').forEach(th => {
            th.style.position = 'sticky';
            th.style.top      = divH + 'px';
            th.style.zIndex   = '15';
            th.style.background = '#6d8196';
        });

        if (headRows[1]) {
            headRows[1].querySelectorAll('th').forEach(th => {
                th.style.position = 'sticky';
                th.style.top      = (divH + firstH) + 'px';
                th.style.zIndex   = '10';
                th.style.background = 'white';
            });
        }

        document.documentElement.style.scrollPaddingTop =
            divH + firstH + secondH + 'px';
    };

    btnFilter.addEventListener('click', () => {
        const hidden = rowFilter.style.display === 'none';
        rowFilter.style.display = hidden ? 'table-row' : 'none';
        rowFilter.style.opacity = hidden ? '1' : '0';
        btnFilter.style.color   = hidden ? '#FFDE21' : 'white';
        updateSticky();
    });

    const showToast = (e, msg) => {
        if (followToast) followToast.remove();

        const t = document.createElement('div');
        t.innerHTML = msg;
        Object.assign(t.style, {
            position: 'fixed',
            background: '#6d8196',
            color: '#fff',
            padding: '8px',
            borderRadius: '3px',
            fontSize: '13px',
            zIndex: 9999,
            pointerEvents: 'none',
            boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
            border: '1px solid rgba(255,255,255,0.1)',
            backdropFilter: 'blur(10px)',
            whiteSpace: 'nowrap'
        });

        document.body.appendChild(t);

        const w = t.offsetWidth, h = t.offsetHeight;
        let x = e.clientX - w - 10;
        let y = e.clientY + 10;

        if (x < 5) x = e.clientX + 15;
        if (y < 5) y = 5;
        if (y + h > innerHeight - 5) y = innerHeight - h - 5;

        t.style.left = x + 'px';
        t.style.top  = y + 'px';

        followToast = t;
    };

    document.querySelectorAll('[data-message]').forEach(el => {
        el.style.cursor = 'help';
        el.addEventListener('mousemove', e => showToast(e, el.dataset.message));
        el.addEventListener('mouseleave', () => {
            if (followToast) followToast.remove();
            followToast = null;
        });
    });

    const markCopyable = () => {
        const cols = {};
        table.querySelectorAll('th[data-column]').forEach((th, i) => {
            const col = th.dataset.column;
            if (col === 'mac' || col === 'serial') cols[col] = i;
        });

        for (const { tr } of rowCache) {
            for (const idx of Object.values(cols)) {
                const cell = tr.children[idx];
                if (!cell) continue;

                const txt = cell.textContent.trim();
                if (!txt) continue;

                cell.style.cursor = 'copy';
                cell.title = 'Click to copy';
                cell.addEventListener('click', () => {
                    navigator.clipboard.writeText(txt).then(() => {
                        const old = cell.textContent;
                        cell.textContent = 'Copied!';
                        cell.style.backgroundColor = '#e8f5e9';
                        setTimeout(() => {
                            cell.textContent = old;
                            cell.style.backgroundColor = '';
                        }, 800);
                    });
                });
            }
        }
    };

    btnExport.addEventListener('click', () => {
        const skip = [14, 15, 16];
        let html = '<html><head><meta charset="UTF-8"><style>td,th{mso-number-format:"\\@"; text-align:"center"; vertical-align:"middle";}</style></head><body><table border="1">';

        const allRows = table.querySelectorAll('tr');
        allRows.forEach((r, i) => {
            if (r.classList?.contains('filter-row')) return;
            if (r.hidden) return;

            html += '<tr>';
            const cells = r.querySelectorAll('td,th');
            cells.forEach((c, idx) => {
                if (!skip.includes(idx)) {
                    const tag = i === 0 ? 'th' : 'td';
                    html += '<' + tag + '>' + c.textContent + '</' + tag + '>';
                }
            });
            html += '</tr>';
        });

        html += '</table></body></html>';

        const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        const url  = URL.createObjectURL(blob);
        const a    = Object.assign(document.createElement('a'), {
            href: url,
            download: 'PsSPM_Report.xls'
        });
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    });

    window.addEventListener('scroll', () => { scrollToTopBtn.classList.toggle('show', window.scrollY > 300); });
    scrollToTopBtn.addEventListener('click', () => { window.scrollTo({ top: 0, behavior: 'smooth' }); });

    window.addEventListener('load', updateSticky);
    window.addEventListener('resize', updateSticky);

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