const API_BASE = "https://digit-excel-taxi.onrender.com";

let selectedFile = null;
let selectedSheet = null;
let processingResult = null;

const fileInput = document.getElementById("policy-file");
const fileNameDisplay = document.getElementById("file-name-display");
const processBtn = document.getElementById("process-button");
const companyInput = document.getElementById("company-name");

// Enable button when file is selected
fileInput.addEventListener("change", () => {
    selectedFile = fileInput.files[0];
    fileNameDisplay.textContent = selectedFile ? selectedFile.name : "Click or drag file here";
    processBtn.disabled = !selectedFile;
});

// Drag & drop support
const dropArea = document.querySelector(".file-upload-label");
dropArea.addEventListener("dragover", e => { e.preventDefault(); dropArea.classList.add("bg-blue-900/30"); });
dropArea.addEventListener("dragleave", () => dropArea.classList.remove("bg-blue-900/30"));
dropArea.addEventListener("drop", e => {
    e.preventDefault();
    dropArea.classList.remove("bg-blue-900/30");
    if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        fileInput.dispatchEvent(new Event("change"));
    }
});

// Tab switching
document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
        document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("border-blue-500", "text-blue-600"));
        document.querySelectorAll(".tab-content").forEach(t => t.classList.add("hidden"));
        btn.classList.add("border-blue-500", "text-blue-600");
        document.getElementById(`tab-${btn.dataset.tab}`).classList.remove("hidden");
    });
});

// Process logic
processBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    // Reset UI
    document.getElementById("initial-message").classList.add("hidden");
    document.getElementById("sheet-selection").classList.add("hidden");
    document.getElementById("processing-spinner").classList.remove("hidden");
    document.getElementById("error-message").classList.add("hidden");
    document.getElementById("results-container").classList.add("hidden");

    try {
        // Step 1: Get sheet names
        const sheetsForm = new FormData();
        sheetsForm.append("file", selectedFile);

        const sheetsRes = await fetch(`${API_BASE}/get-sheets`, {
            method: "POST",
            body: sheetsForm
        });

        const sheetsData = await sheetsRes.json();

        if (!sheetsData.success) {
            throw new Error(sheetsData.message || "Could not read sheets");
        }

        if (sheetsData.total_sheets === 1 || sheetsData.sheets.length === 1) {
            // Auto process single sheet
            await processWithSheet("");
        } else {
            // Show sheet selection
            showSheetSelection(sheetsData.sheets);
        }
    } catch (err) {
        showError(err.message || "Failed to connect to backend");
    }
});

async function processWithSheet(sheetName) {
    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("company_name", companyInput.value.trim());
    if (sheetName) formData.append("sheet_name", sheetName);

    try {
        const res = await fetch(`${API_BASE}/taxi`, {
            method: "POST",
            body: formData
        });

        const data = await res.json();

        document.getElementById("processing-spinner").classList.add("hidden");

        if (!data.success) {
            throw new Error(data.error || data.message || "Processing failed");
        }

        processingResult = data;
        renderResults(data);
    } catch (err) {
        showError(err.message);
    }
}

function showSheetSelection(sheets) {
    const list = document.getElementById("sheet-list");
    list.innerHTML = "";

    sheets.forEach(name => {
        const div = document.createElement("div");
        div.className = "sheet-option p-4 bg-gray-100 rounded-lg cursor-pointer border border-gray-300";
        div.textContent = name;
        div.onclick = () => {
            list.querySelectorAll(".sheet-option").forEach(el => el.classList.remove("selected"));
            div.classList.add("selected");
            selectedSheet = name;
            document.getElementById("confirm-sheet-btn").disabled = false;
        };
        list.appendChild(div);
    });

    document.getElementById("sheet-selection").classList.remove("hidden");
    document.getElementById("processing-spinner").classList.add("hidden");
}

document.getElementById("confirm-sheet-btn").onclick = () => {
    if (selectedSheet) {
        processWithSheet(selectedSheet);
    }
};

document.getElementById("cancel-sheet-btn").onclick = () => {
    document.getElementById("sheet-selection").classList.add("hidden");
    document.getElementById("initial-message").classList.remove("hidden");
};

function renderResults(data) {
    document.getElementById("results-container").classList.remove("hidden");
    document.getElementById("company-name-display").textContent = data.company_name || "Digit";
    document.getElementById("total-records").textContent = data.total_records || 0;
    document.getElementById("avg-payin").textContent = data.avg_payin ? `${data.avg_payin}%` : "0.0%";
    document.getElementById("unique-segments").textContent = data.unique_segments || 0;

    // Table
    const tbody = document.getElementById("results-table-body");
    tbody.innerHTML = "";
    (data.calculated_data || []).forEach(r => {
        const tr = document.createElement("tr");
        tr.className = "hover:bg-blue-50 transition-colors";
        tr.innerHTML = `
            <td class="p-4 border-b">${r.State || "-"}</td>
            <td class="p-4 border-b">${r["Location/Cluster"] || "-"}</td>
            <td class="p-4 border-b text-gray-700">${r["Mapped Segment"] || "-"}</td>
            <td class="p-4 border-b font-bold text-blue-700">${r["Payin (CD2)"] || "0%"}</td>
            <td class="p-4 border-b font-bold text-green-700">${r["Calculated Payout"] || "0%"}</td>
            <td class="p-4 border-b text-gray-600 text-xs">${r["Formula Used"] || "-"}</td>
            <td class="p-4 border-b text-gray-500 text-xs">${r["Rule Explanation"] || "-"}</td>
        `;
        tbody.appendChild(tr);
    });

    // Detection
    const patternDiv = document.getElementById("pattern-summary");
    patternDiv.innerHTML = "";
    Object.entries(data.patterns_detected || {}).forEach(([k, v]) => {
        const item = document.createElement("div");
        item.className = "flex justify-between p-3 bg-gray-50 rounded border";
        item.innerHTML = `<span class="font-medium">${k.toUpperCase()}</span><span class="font-bold text-blue-600">${v}</span>`;
        patternDiv.appendChild(item);
    });

    const procList = document.getElementById("processors-list");
    procList.innerHTML = "";
    (data.processors_used || []).forEach(p => {
        const li = document.createElement("li");
        li.className = "bg-blue-50 p-3 rounded flex items-center gap-2";
        li.innerHTML = `<i class="fas fa-cog text-blue-500"></i> ${p}`;
        procList.appendChild(li);
    });

    // Raw
    document.getElementById("parsed-data").textContent = JSON.stringify(data.calculated_data, null, 2);

    // Downloads
    document.getElementById("btn-excel").onclick = () => {
        const a = document.createElement("a");
        a.href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${data.excel_data}`;
        a.download = `TAXI_${data.company_name || "results"}_${new Date().toISOString().slice(0,10)}.xlsx`;
        a.click();
    };

    document.getElementById("btn-csv").onclick = () => {
        const blob = new Blob([data.csv_data], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `TAXI_${data.company_name || "results"}_${new Date().toISOString().slice(0,10)}.csv`;
        a.click();
        URL.revokeObjectURL(url);
    };

    document.getElementById("btn-json").onclick = () => {
        const blob = new Blob([JSON.stringify(data.calculated_data, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `TAXI_${data.company_name || "results"}_${new Date().toISOString().slice(0,10)}.json`;
        a.click();
        URL.revokeObjectURL(url);
    };
}

function showError(msg) {
    document.getElementById("error-text").textContent = msg;
    document.getElementById("error-message").classList.remove("hidden");
    document.getElementById("processing-spinner").classList.add("hidden");
}
