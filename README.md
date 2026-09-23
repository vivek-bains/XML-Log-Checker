<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power Platform Dependency Dashboard</title>

    <style>
        * { box-sizing: border-box; }

        body {
            max-width: 1500px;
            margin: 0 auto;
            padding: 28px;
            background: #f3f4f6;
            color: #1f2937;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        h1 { margin: 0 0 8px; color: #111827; }
        p { margin: 0 0 20px; color: #4b5563; }

        textarea {
            display: block;
            width: 100%;
            height: 150px;
            padding: 12px;
            margin-bottom: 14px;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            font: 13px Consolas, monospace;
            resize: vertical;
        }

        .actions-bar {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 18px;
        }

        button {
            padding: 10px 16px;
            border: 0;
            border-radius: 7px;
            background: #2563eb;
            color: white;
            font: inherit;
            font-weight: 600;
            cursor: pointer;
        }

        button:hover { background: #1d4ed8; }

        .btn-export {
            display: none;
            background: #059669;
        }

        .btn-export:hover { background: #047857; }

        .btn-email {
            display: none;
            background: #4f46e5;
        }

        .btn-email:hover { background: #4338ca; }

        .counter-badge {
            display: none;
            width: fit-content;
            margin-bottom: 14px;
            padding: 9px 15px;
            border-radius: 30px;
            background: #e5e7eb;
            font-weight: 600;
        }

        #totalCount { margin-left: 5px; color: #2563eb; }

        .error-msg {
            display: none;
            margin-bottom: 14px;
            color: #b91c1c;
            font-weight: 600;
        }

        .table-container {
            display: none;
            width: 100%;
            overflow-x: auto;
            border: 1px solid #e5e7eb;
            border-radius: 9px;
            background: white;
            box-shadow: 0 1px 3px rgb(0 0 0 / 8%);
        }

        table {
            table-layout: fixed;
            border-collapse: collapse;
            font-size: 14px;
        }

        th, td {
            padding: 10px 12px;
            border-bottom: 1px solid #e5e7eb;
            text-align: left;
            vertical-align: top;
            line-height: 1.4;
            overflow-wrap: anywhere;
        }

        th {
            position: relative;
            background: #f9fafb;
            color: #374151;
            font-weight: 700;
            white-space: nowrap;
            user-select: none;
        }

        .sort-button {
            width: 100%;
            padding: 0 13px 0 0;
            border: 0;
            background: none;
            color: inherit;
            text-align: left;
            font: inherit;
            cursor: pointer;
        }

        .sort-button:hover {
            background: none;
            color: #1d4ed8;
        }

        .sort-indicator {
            margin-left: 5px;
            color: #2563eb;
        }

        tbody tr:hover { background: #f8fafc; }
        tbody tr:last-child td { border-bottom: 0; }

        .resize-handle {
            position: absolute;
            top: 0;
            right: 0;
            width: 9px;
            height: 100%;
            cursor: col-resize;
            touch-action: none;
        }

        .resize-handle:hover,
        .resize-handle.active {
            background: #93c5fd;
        }

        .hint {
            display: none;
            margin: 12px 0 0;
            color: #6b7280;
            font-size: 13px;
        }

        @media (max-width: 700px) {
            body { padding: 16px; }
        }
    </style>
</head>

<body>
    <h1>Power Platform Dependency Dashboard</h1>
    <p>Paste your D365 solution import log below to view its missing dependencies.</p>

    <textarea
        id="logInput"
        placeholder="Paste your D365 JSON log or MissingDependencies XML here..."
    ></textarea>

    <div class="actions-bar">
        <button type="button" onclick="cleanLog()">Parse &amp; Process Log</button>
        <button id="btnCsv" class="btn-export" type="button" onclick="exportData('csv')">
            Export CSV
        </button>
        <button id="btnExcel" class="btn-export" type="button" onclick="exportData('excel')">
            Export for Excel
        </button>
        <button id="btnEmail" class="btn-email" type="button" onclick="shareViaEmail()">
            Share Summary via Email
        </button>

        <button type="button" onclick="clearAll()" style="background: #6b7280;">
    Clear All
</button>

    </div>

    <div id="counterArea" class="counter-badge">
        Total Missing Dependencies: <span id="totalCount">0</span>
    </div>

    <div id="errorMsg" class="error-msg"></div>

    <div id="tableContainer" class="table-container">
        <table id="dependencyTable">
            <colgroup>
                <col style="width: 55px">
                <col style="width: 240px">
                <col style="width: 230px">
                <col style="width: 210px">
                <col style="width: 110px">
                <col style="width: 250px">
                <col style="width: 150px">
                <col style="width: 355px">
            </colgroup>

            <thead>
                <tr>
                    <th data-key="number"><button class="sort-button" type="button">#</button></th>
                    <th data-key="displayName"><button class="sort-button" type="button">Display Name</button></th>
                    <th data-key="nameOrId"><button class="sort-button" type="button">Name/ID</button></th>
                    <th data-key="tableName"><button class="sort-button" type="button">Table</button></th>
                    <th data-key="requiredType"><button class="sort-button" type="button">Type</button></th>
                    <th data-key="requiredBy"><button class="sort-button" type="button">Required By</button></th>
                    <th data-key="dependentType"><button class="sort-button" type="button">Type</button></th>
                    <th data-key="action"><button class="sort-button" type="button">Action</button></th>
                </tr>
            </thead>

            <tbody id="tableBody"></tbody>
        </table>
    </div>

    <p id="tableHint" class="hint">
        Click a heading to sort A–Z or Z–A. Drag its right edge to resize the column.
    </p>

    <script>
        "use strict";

        let dependencies = [];
        let sortKey = "number";
        let sortDirection = "asc";

        const tableColumns = [
            { key: "number", label: "#" },
            { key: "displayName", label: "Display Name" },
            { key: "nameOrId", label: "Name/ID" },
            { key: "tableName", label: "Table" },
            { key: "requiredType", label: "Type" },
            { key: "requiredBy", label: "Required By" },
            { key: "dependentType", label: "Dependent Type" },
            { key: "action", label: "Action" }
        ];

        function componentType(code) {
            const types = {
                "26": "View",
                "60": "Form / Dashboard"
            };

            return types[code] || (code ? "Type " + code : "Unknown");
        }

        function showError(message) {
            const error = document.getElementById("errorMsg");
            error.textContent = message;
            error.style.display = "block";
        }

        function resetResults() {
            dependencies = [];
            sortKey = "number";
            sortDirection = "asc";

            document.getElementById("tableBody").replaceChildren();
            document.getElementById("tableContainer").style.display = "none";
            document.getElementById("counterArea").style.display = "none";
            document.getElementById("errorMsg").style.display = "none";
            document.getElementById("tableHint").style.display = "none";

            ["btnCsv", "btnExcel", "btnEmail"].forEach(id => {
                document.getElementById(id).style.display = "none";
            });

            updateSortHeadings();
        }

        function cleanLog() {
            resetResults();

            const input = document.getElementById("logInput").value.trim();

            if (!input) {
                showError("Paste your D365 solution log first.");
                return;
            }

            try {
                let messages;

                try {
                    const log = JSON.parse(input);
                    const entries = Array.isArray(log) ? log : [log];

                    messages = entries.map(entry =>
                        String(entry.Message || entry.message || "")
                    );
                } catch {
                    messages = [input];
                }

                for (const message of messages) {
                    const xmlMatch = message.match(
                        /<MissingDependencies\b[\s\S]*?<\/MissingDependencies>/i
                    );

                    if (!xmlMatch) continue;

                    const xml = new DOMParser().parseFromString(
                        xmlMatch[0],
                        "application/xml"
                    );

                    if (xml.getElementsByTagName("parsererror").length) {
                        throw new Error(
                            "The dependency XML inside the log could not be read."
                        );
                    }

                    const items = xml.getElementsByTagName("MissingDependency");

                    for (const item of items) {
                        const required =
                            item.getElementsByTagName("Required")[0];

                        const dependent =
                            item.getElementsByTagName("Dependent")[0];

                        if (!required) continue;

                        const tableName = [
                            required.getAttribute("parentDisplayName"),
                            required.getAttribute("parentSchemaName")
                        ].filter(Boolean).join(" / ") || "Not specified";

                        dependencies.push({
                            number: dependencies.length + 1,

                            displayName:
                                required.getAttribute("displayName") || "",

                            nameOrId:
                                required.getAttribute("schemaName") ||
                                required.getAttribute("id") ||
                                "",

                            tableName,

                            requiredType: componentType(
                                required.getAttribute("type")
                            ),

                            requiredBy:
                                dependent?.getAttribute("displayName") ||
                                dependent?.getAttribute("id") ||
                                "",

                            dependentType: componentType(
                                dependent?.getAttribute("type")
                            ),

                            action:
                                "Add the missing component to the target " +
                                "environment, or remove its reference " +
                                "from the dependent component."
                        });
                    }
                }

                if (!dependencies.length) {
                    showError(
                        "No MissingDependency entries were found in this log."
                    );
                    return;
                }

                renderTable();

                document.getElementById("totalCount").textContent =
                    dependencies.length;

                document.getElementById("counterArea").style.display =
                    "inline-flex";

                document.getElementById("tableContainer").style.display =
                    "block";

                document.getElementById("tableHint").style.display =
                    "block";

                ["btnCsv", "btnExcel", "btnEmail"].forEach(id => {
                    document.getElementById(id).style.display =
                        "inline-block";
                });

            } catch (error) {
                resetResults();
                showError(error.message);
            }
        }

        function getSortedDependencies() {
            const sorted = [...dependencies];

            sorted.sort((a, b) => {
                let result;

                if (sortKey === "number") {
                    result = a.number - b.number;
                } else {
                    result = String(a[sortKey] || "").localeCompare(
                        String(b[sortKey] || ""),
                        undefined,
                        {
                            sensitivity: "base",
                            numeric: true
                        }
                    );

                    // Keep the original order when values are equal.
                    if (result === 0) {
                        result = a.number - b.number;
                    }
                }

                return sortDirection === "asc" ? result : -result;
            });

            return sorted;
        }

        function renderTable() {
            const body = document.getElementById("tableBody");
            const fragment = document.createDocumentFragment();

            body.replaceChildren();

            for (const item of getSortedDependencies()) {
                const row = document.createElement("tr");

                for (const column of tableColumns) {
                    const cell = document.createElement("td");

                    // Log values are displayed as text, never as HTML.
                    cell.textContent = item[column.key] ?? "";

                    row.appendChild(cell);
                }

                fragment.appendChild(row);
            }

            body.appendChild(fragment);
            updateSortHeadings();
        }

        function sortBy(key) {
            if (!dependencies.length) return;

            if (sortKey === key) {
                sortDirection =
                    sortDirection === "asc" ? "desc" : "asc";
            } else {
                sortKey = key;
                sortDirection = "asc";
            }

            renderTable();
        }

        function updateSortHeadings() {
            document
                .querySelectorAll("#dependencyTable thead th")
                .forEach(header => {
                    const button = header.querySelector(".sort-button");
                    const key = header.dataset.key;
                    const label = tableColumns.find(
                        column => column.key === key
                    ).label;

                    button.textContent =
                        label +
                        (key === sortKey
                            ? (sortDirection === "asc" ? " ▲" : " ▼")
                            : "");

                    header.setAttribute(
                        "aria-sort",
                        key === sortKey
                            ? (sortDirection === "asc"
                                ? "ascending"
                                : "descending")
                            : "none"
                    );
                });
        }

        function exportData(format) {
            if (!dependencies.length) return;

            const headers = tableColumns.map(column => column.label);

            // Export rows in the same order currently shown on screen.
            const rows = getSortedDependencies().map(item =>
                tableColumns.map(column => item[column.key] ?? "")
            );

            const csvCell = value =>
                '"' + String(value).replace(/"/g, '""') + '"';

            const csv = [headers, ...rows]
                .map(row => row.map(csvCell).join(","))
                .join("\r\n");

            const blob = new Blob(["\uFEFF" + csv], {
                type: "text/csv;charset=utf-8"
            });

            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");

            link.href = url;
            link.download = format === "excel"
                ? "missing_dependencies_excel.csv"
                : "missing_dependencies.csv";

            document.body.appendChild(link);
            link.click();
            link.remove();

            setTimeout(() => URL.revokeObjectURL(url), 1000);
        }

        function shareViaEmail() {
            if (!dependencies.length) return;

            const lines = getSortedDependencies().map(item =>
                `${item.number}. ${item.displayName} ` +
                `(${item.requiredType})\n` +
                `Name/ID: ${item.nameOrId}\n` +
                `Table: ${item.tableName}\n` +
                `Required by: ${item.requiredBy} ` +
                `(${item.dependentType})`
            );

            const subject = encodeURIComponent(
                "D365 Missing Dependencies"
            );

            const body = encodeURIComponent(
                `Missing dependencies: ${dependencies.length}\n\n` +
                lines.join("\n\n")
            );

            window.location.href =
                `mailto:?subject=${subject}&body=${body}`;
        }

        function updateTableWidth() {
            const table = document.getElementById("dependencyTable");
            const columns = table.querySelectorAll("colgroup col");

            const totalWidth = Array.from(columns).reduce(
                (total, column) =>
                    total + parseFloat(column.style.width || "0"),
                0
            );

            table.style.width = totalWidth + "px";
        }

        function enableTableControls() {
            const table = document.getElementById("dependencyTable");
            const headers = table.querySelectorAll("thead th");
            const columns = table.querySelectorAll("colgroup col");

            headers.forEach((header, index) => {
                header
                    .querySelector(".sort-button")
                    .addEventListener("click", () => {
                        sortBy(header.dataset.key);
                    });

                const handle = document.createElement("span");

                handle.className = "resize-handle";
                handle.title = "Drag to resize column";
                header.appendChild(handle);

                handle.addEventListener("pointerdown", event => {
                    event.preventDefault();
                    event.stopPropagation();

                    const startX = event.clientX;
                    const startWidth =
                        columns[index].getBoundingClientRect().width;

                    handle.classList.add("active");
                    handle.setPointerCapture(event.pointerId);

                    function onMove(moveEvent) {
                        const newWidth = Math.max(
                            70,
                            startWidth + moveEvent.clientX - startX
                        );

                        columns[index].style.width = newWidth + "px";
                        updateTableWidth();
                    }

                    function onEnd() {
                        handle.classList.remove("active");
                        handle.removeEventListener("pointermove", onMove);
                        handle.removeEventListener("pointerup", onEnd);
                        handle.removeEventListener(
                            "pointercancel",
                            onEnd
                        );
                    }

                    handle.addEventListener("pointermove", onMove);
                    handle.addEventListener("pointerup", onEnd);
                    handle.addEventListener("pointercancel", onEnd);
                });
            });
        }

function clearAll() {
    document.getElementById("logInput").value = "";
    resetResults();
    document.getElementById("totalCount").textContent = "0";
    document.getElementById("logInput").focus();
}



        enableTableControls();
        updateTableWidth();
        updateSortHeadings();
    </script>
</body>
</html>
