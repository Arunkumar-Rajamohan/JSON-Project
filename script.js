document.getElementById('excelFile').addEventListener('change', handleFileUpload);
document.getElementById('searchInput').addEventListener('input', handleSearch);

let currentData = [];
let filteredData = [];
let currentPage = 1;
const rowsPerPage = 10;

function handleFileUpload(event) {
    const file = event.target.files[0];
    const message = document.getElementById('message');
    const sheetSelectorContainer = document.getElementById('sheetSelectorContainer');
    const sheetSelector = document.getElementById('sheetSelector');

    if (!file) {
        alert("Please select an Excel file.");
        return;
    }

    message.style.display = 'block';
    message.textContent = "Processing your file...";

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        sheetSelector.innerHTML = '';
        workbook.SheetNames.forEach((sheetName, index) => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            sheetSelector.appendChild(option);
        });

        sheetSelectorContainer.style.display = 'block';
        sheetSelector.addEventListener('change', () => {
            const selectedSheet = sheetSelector.value;
            const sheet = workbook.Sheets[selectedSheet];
            currentData = XLSX.utils.sheet_to_json(sheet);
            if (currentData.length === 0) {
                message.textContent = "The selected sheet is empty.";
                return;
            }
            message.style.display = 'none';
            filteredData = [...currentData];
            currentPage = 1;
            displayTable();
        });

        sheetSelector.dispatchEvent(new Event('change'));
    };
    reader.readAsArrayBuffer(file);
}

function displayTable() {
    const table = document.getElementById('dataTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const paginationControls = document.getElementById('paginationControls');
    const downloadJSON = document.getElementById('downloadJSON');

    thead.innerHTML = '';
    tbody.innerHTML = '';
    paginationControls.innerHTML = '';

    if (filteredData.length === 0) {
        thead.innerHTML = '<tr><th>No data available</th></tr>';
        return;
    }

    const headers = Object.keys(filteredData[0]);
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    const start = (currentPage - 1) * rowsPerPage;
    const end = start + rowsPerPage;
    const pageData = filteredData.slice(start, end);

    pageData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    for (let i = 1; i <= totalPages; i++) {
        const button = document.createElement('button');
        button.textContent = i;
        button.className = `btn btn-sm ${i === currentPage ? 'btn-primary' : 'btn-light'}`;
        button.addEventListener('click', () => {
            currentPage = i;
            displayTable();
        });
        paginationControls.appendChild(button);
    }

    downloadJSON.style.display = 'block';
    downloadJSON.onclick = () => {
        const jsonBlob = new Blob([JSON.stringify(filteredData, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(jsonBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'data.json';
        a.click();
        URL.revokeObjectURL(url);
    };
}

function handleSearch(event) {
    const query = event.target.value.toLowerCase();
    filteredData = currentData.filter(row =>
        Object.values(row).some(value => String(value).toLowerCase().includes(query))
    );
    currentPage = 1;
    displayTable();
}
