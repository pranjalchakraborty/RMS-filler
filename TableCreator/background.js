// A single, robust function to handle all Excel download requests
function handleDownload(data, filename) {
    if (!data || data.length === 0) {
        console.error(`[Background Script] Download request for "${filename}" received, but no data was provided.`);
        return;
    }
    try {
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Routine");

        // Auto-size columns for better readability
        const colWidths = Object.keys(data[0]).map(key => ({
            wch: Math.max(...data.map(item => (item[key] || "").toString().length)) + 2
        }));
        worksheet['!cols'] = colWidths;

        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);

        // Use saveAs: true to prompt the user where to save the file
        chrome.downloads.download({ url: url, filename: filename, saveAs: true }, () => {
            setTimeout(() => URL.revokeObjectURL(url), 1000); // Clean up memory
        });
    } catch (e) {
        console.error(`[Background Script] FATAL ERROR creating Excel file "${filename}":`, e);
    }
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    switch (request.action) {
        case 'download_data':
            console.log("[Background Script] Handling 'download_data' request.");
            handleDownload(request.data, 'routine_scraped.xlsx');
            break;

        case 'download_updated_excel':
            console.log("[Background Script] Handling 'download_updated_excel' request.");
            handleDownload(request.data, 'routine_filled_report.xlsx');
            break;

        case 'parse_excel_data':
            console.log("[Background Script] Handling 'parse_excel_data' request.");
            try {
                const base64Data = request.data.substring(request.data.indexOf(',') + 1);
                const workbook = XLSX.read(base64Data, { type: 'base64' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: "" });

                chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
                    if (tabs[0]) {
                        chrome.tabs.sendMessage(tabs[0].id, { action: 'fill_form_data', data: jsonData });
                    }
                });
            } catch (e) {
                console.error("[Background Script] FATAL ERROR parsing Excel data:", e);
            }
            break;
    }
    
    sendResponse({ status: 'completed' });
    return true;
});