// The first line MUST be importScripts in a Manifest V3 service worker.
try {
  importScripts('xlsx.full.min.js');
} catch (e) {
  console.error("Failed to load the XLSX library. Make sure xlsx.full.min.js is in the TableCreator directory.", e);
}

function handleDownload(data, filename) {
    if (!data || data.length === 0) {
        console.error(`[Background] Download for "${filename}" failed: No data.`);
        return;
    }
    try {
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Routine");
        const colWidths = Object.keys(data[0]).map(key => ({ wch: Math.max(...data.map(item => (item[key] || "").toString().length)) + 2 }));
        worksheet['!cols'] = colWidths;
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const reader = new FileReader();
        reader.onload = () => {
            const dataUrl = reader.result;
            chrome.downloads.download({
                url: dataUrl,
                filename: filename,
                saveAs: true
            });
        };
        reader.readAsDataURL(blob);

    } catch (e) {
        console.error(`[Background] Error creating Excel file "${filename}":`, e);
    }
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    switch (request.action) {
        case 'download_data':
            handleDownload(request.data, 'routine_scraped.xlsx');
            break;
        case 'download_updated_excel':
            handleDownload(request.data, 'routine_filled_report.xlsx');
            break;
        case 'parse_excel_data':
            try {
                const base64Data = request.data.substring(request.data.indexOf(',') + 1);
                const workbook = XLSX.read(base64Data, { type: 'base64' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: "" });
                
                chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
                    if (tabs[0] && tabs[0].id) {
                        chrome.scripting.executeScript({
                            target: { tabId: tabs[0].id },
                            files: ['content.js']
                        }).then(() => {
                            chrome.tabs.sendMessage(tabs[0].id, { action: 'fill_form_data', data: jsonData });
                        }).catch(err => console.error("Background script failed to inject content script:", err));
                    }
                });
            } catch (e) {
                console.error("[Background] Error parsing Excel data:", e);
            }
            break;
    }
    sendResponse({ status: 'completed' });
    return true;
});