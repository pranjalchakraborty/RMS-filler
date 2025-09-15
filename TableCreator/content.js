(function() {
    'use strict';

    // ===================================================================================
    // SECTION 1: CORE UTILITY FUNCTIONS (Used by both Scraper and Filler)
    // ===================================================================================

    const waitForElement = (selector, context = document, timeout = 7000) => {
        return new Promise(resolve => {
            const startTime = Date.now();
            const interval = setInterval(() => {
                const element = context.querySelector(selector);
                if (element) {
                    clearInterval(interval);
                    resolve(element);
                } else if (Date.now() - startTime > timeout) {
                    clearInterval(interval);
                    console.error(`Timeout: Element "${selector}" not found.`);
                    resolve(null);
                }
            }, 100);
        });
    };

    const clickAndWait = async (element, delay = 3000) => {
        if (element) {
            element.click();
            await new Promise(resolve => setTimeout(resolve, delay));
        }
    };

    const selectOptionByText = (selectElement, textToFind) => {
        if (!selectElement || !textToFind) return false;
        const text = String(textToFind).trim().toLowerCase();
        const option = [...selectElement.options].find(opt => opt.text.trim().toLowerCase() === text);
        if (option) {
            selectElement.value = option.value;
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        return false;
    };


    // ===================================================================================
    // SECTION 2: DATA SCRAPING LOGIC (Restored and Complete)
    // ===================================================================================

    const showCalendarProgrammatically = () => {
        return new Promise(resolve => {
            const scriptContent = `$('#myResultbacknew.in .datepicker_with_range').datepicker('show');`;
            const script = document.createElement('script');
            script.textContent = scriptContent;
            document.body.appendChild(script);
            document.body.removeChild(script);
            setTimeout(resolve, 300);
        });
    };

    const scrapeCalendar = async () => {
        const dates = new Set();
        let calendar = await waitForElement('.datepicker.datepicker-dropdown');
        if (!calendar) return [];
    
        while (true) {
            calendar.querySelectorAll('.day.allowed-date:not(.old):not(.new)').forEach(el => {
                const timestamp = el.getAttribute('data-date');
                if (timestamp) {
                    const date = new Date(parseInt(timestamp));
                    dates.add(date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2));
                }
            });
            const prevBtn = calendar.querySelector('.prev:not(.disabled)');
            if (prevBtn) {
                await clickAndWait(prevBtn, 500);
                calendar = await waitForElement('.datepicker.datepicker-dropdown');
                if (!calendar) break;
            } else {
                break;
            }
        }
        return Array.from(dates);
    };

    const scrapeRoutine = async () => {
        console.log("Scraping process initiated...");
        alert("Starting to scrape routine data. Please wait.");
        const data = [];
        const buttons = document.querySelectorAll('a.class_execution_data');
    
        for (const button of buttons) {
            try {
                const cell = button.closest('td');
                const row = button.closest('tr');
                if (!row || !cell) continue;

                const columnNumber = cell.cellIndex;
                const day = row.cells[0]?.textContent.trim();

                await clickAndWait(button);
                const modal = await waitForElement('#myResultbacknew.in');
                if (!modal) continue;

                const subjectSelect = await waitForElement('#subject_name_no', modal);
                const subject = (subjectSelect && subjectSelect.options.length > 1) ? subjectSelect.options[1].text.trim() : 'Subject Not Found';

                document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                await showCalendarProgrammatically();
                const dates = await scrapeCalendar();
                
                dates.forEach(date => {
                    data.push({ "Days": day, "Column": columnNumber, "Subject": subject, "Date": date });
                });

                await clickAndWait(modal.querySelector('button.close_btn'), 500);
            } catch (error) {
                console.error("An error occurred during scraping a routine slot:", error);
                const openModal = document.querySelector('#myResultbacknew.in');
                if (openModal) {
                    await clickAndWait(openModal.querySelector('button.close_btn'), 500);
                }
            }
        }
    
        if (data.length > 0) {
            alert(`Scraping complete! Found ${data.length} records. The file download will now begin.`);
            chrome.runtime.sendMessage({ action: 'download_data', data: data });
        } else {
            alert("Scraping finished, but no data was found. Please check the console for errors.");
        }
    };


    // ===================================================================================
    // SECTION 3: FORM FILLING LOGIC (Confirmed Working)
    // ===================================================================================
    
    const normalizeDataKeys = (data) => data.map(row => Object.keys(row).reduce((acc, key) => { acc[key.trim()] = row[key]; return acc; }, {}));

    const fillFormsFromExcel = async (data) => {
        let processedData = normalizeDataKeys(data);
        const routineTable = document.querySelector('.table.table-bordered');
        if (!routineTable) { alert("Fatal Error: Could not find routine table."); return; }

        const rowsToProcess = processedData.filter(row => String(row['Submitted'] || '').trim() === '' && ['yes', 'no'].includes(String(row['Class Execution'] || '').trim().toLowerCase()));

        if (rowsToProcess.length === 0) { alert("No new records to process were found in the Excel file."); return; }
        
        alert(`Found ${rowsToProcess.length} records to process. The automation will now begin.`);

        for (const rowData of rowsToProcess) {
            console.group(`Processing Excel Row for Date: ${rowData.Date}`);
            try {
                const tableRow = [...routineTable.querySelectorAll('tbody > tr')].find(r => r.cells[0]?.textContent.trim().toLowerCase() === rowData.Days.toLowerCase());
                if (!tableRow) throw new Error(`Could not find table row for day: ${rowData.Days}`);

                const execButton = tableRow.cells[parseInt(rowData.Column, 10)]?.querySelector('a.class_execution_data');
                if (!execButton) throw new Error(`Could not find 'Class Execution' button`);

                await clickAndWait(execButton, 5000);
                const modal = await waitForElement('#myResultbacknew.in');
                if (!modal) throw new Error("Execution modal did not open.");

                (await waitForElement('input.datepicker_with_range', modal)).value = rowData.Date;
                const classExecDropdown = await waitForElement('#class_execution_yes_no', modal);
                const executionStatus = String(rowData['Class Execution']).trim().toLowerCase();

                if (executionStatus === 'yes') {
                    classExecDropdown.value = '1';
                    classExecDropdown.dispatchEvent(new Event('change', { bubbles: true }));
                    await new Promise(r => setTimeout(r, 2000)); 
                    if (!selectOptionByText(await waitForElement('#subject_name_no', modal), rowData.Subject)) throw new Error(`Subject "${rowData.Subject}" could not be selected.`);
                    (await waitForElement('#class_execution_yes', modal)).value = rowData['Topic (Yes)'] || '';
                    (await waitForElement('#registerd_class', modal)).value = rowData['Total Students'] || '';
                    (await waitForElement('#attended_class', modal)).value = rowData['Attended Students'] || '';
                } else if (executionStatus === 'no') {
                    classExecDropdown.value = '0';
                    classExecDropdown.dispatchEvent(new Event('change', { bubbles: true }));
                    await new Promise(r => setTimeout(r, 2000));
                    if (!selectOptionByText(await waitForElement('#class_execution_no', modal), rowData['Reason (No)'])) throw new Error(`Reason "${rowData['Reason (No)']}" could not be selected.`);
                }

                //if (window.confirm(`Ready to submit:\n\nDay: ${rowData.Days}, Date: ${rowData.Date}\nStatus: ${rowData['Class Execution']}\n\nClick OK to submit.`)) {
                console.log("Confirmation step skipped. Submitting form automatically...");    
                await clickAndWait(modal.querySelector('input.confirm_result_revart_back'), 5000);
                    rowData.Submitted = 'Success';
                //} else {
                    //await clickAndWait(modal.querySelector('button.close_btn'), 1000);
                    //rowData.Submitted = 'Cancelled by User';
                //}
            } catch (error) {
                console.error('An error occurred processing row:', error);
                rowData.Submitted = `Error: ${error.message}`;
                const openModal = document.querySelector('#myResultbacknew.in');
                if (openModal) await clickAndWait(openModal.querySelector('button.close_btn'), 500);
            } finally {
                console.groupEnd();
            }
        }
        
        alert("Finished processing all entries. The updated Excel report will now be downloaded.");
        chrome.runtime.sendMessage({ action: 'download_updated_excel', data: processedData });
    };

    // ===================================================================================
    // SECTION 4: MAIN MESSAGE LISTENER
    // ===================================================================================

    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === 'scrape') {
            scrapeRoutine();
        } else if (request.action === 'fill_form_data') {
            fillFormsFromExcel(request.data);
        }
        sendResponse({ status: 'action received' });
        return true; 
    });

})();