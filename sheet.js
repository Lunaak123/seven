function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

const sheetName = getQueryParam('sheetName');
const fileUrl = getQueryParam('fileUrl');

(async () => {
    if (!fileUrl || !sheetName) {
        alert("Invalid sheet data.");
        return;
    }

    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            alert("Sheet not found.");
            return;
        }

        // Display the Excel sheet content in HTML table format
        const sheetContentDiv = document.getElementById('sheet-content');
        const sheetHtml = XLSX.utils.sheet_to_html(sheet, { id: "excel-sheet", editable: false });
        sheetContentDiv.innerHTML = sheetHtml;

        // Make the table fill the page like Excel
        const table = document.getElementById('excel-sheet');
        table.style.width = '100%'; // Make it take full page width
        table.style.tableLayout = 'fixed'; // Equal width for cells
    } catch (error) {
        console.error("Error displaying the Excel sheet:", error);
        alert("Failed to load the Excel sheet. Please try again.");
    }
})();
