(function() {
    'use strict';

    // ===================================================================================
    // SECTION 1: CORE UTILITY FUNCTIONS (Defined Once)
    // ===================================================================================

    const pollForCondition = (condition, timeout = 45000, timeoutMessage = 'Polling timed out') => {
        return new Promise(resolve => {
            const startTime = Date.now();
            const interval = setInterval(() => {
                if (condition()) {
                    clearInterval(interval);
                    resolve(true);
                } else if (Date.now() - startTime > timeout) {
                    clearInterval(interval);
                    console.error(`Timeout: ${timeoutMessage}`);
                    resolve(false);
                }
            }, 50);
        });
    };

    const clickThenPoll = async (elementToClick, condition, timeoutMessage) => {
        if (!elementToClick) {
            console.error("Attempted to click a null element.");
            return false;
        }
        elementToClick.click();
        return await pollForCondition(condition, 45000, timeoutMessage);
    };

    const waitForElement = (selector, context = document, timeout = 20000) => {
        return new Promise(resolve => {
            const startTime = Date.now();
            const interval = setInterval(() => {
                const element = context.querySelector(selector);
                if (element) {
                    clearInterval(interval);
                    resolve(element);
                } else if (Date.now() - startTime > timeout) {
                    clearInterval(interval);
                    console.error(`Timeout: Element "${selector}" not found after 20 seconds.`);
                    resolve(null);
                }
            }, 100);
        });
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
    // SECTION 2: DATA SCRAPING LOGIC
    // ===================================================================================

    const scrapeCalendarInDirection = async (dateSet, direction) => {
        const calendarExists = await pollForCondition(() => document.querySelector('.datepicker.datepicker-dropdown'), 5000, "Calendar did not appear.");
        if (!calendarExists) return;

        let calendar = document.querySelector('.datepicker.datepicker-dropdown');
        while (true) {
            calendar.querySelectorAll('.day.allowed-date:not(.old):not(.new)').forEach(el => {
                const timestamp = el.getAttribute('data-date');
                if (timestamp) {
                    const date = new Date(parseInt(timestamp));
                    dateSet.add(date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2));
                }
            });
            const navButton = calendar.querySelector(`.${direction}:not(.disabled)`);
            if (navButton) {
                navButton.click();
                await new Promise(r => setTimeout(r, 500));
                calendar = document.querySelector('.datepicker.datepicker-dropdown');
            } else {
                break;
            }
        }
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

                const modalAppeared = await clickThenPoll(button, () => document.querySelector('#myResultbacknew'), "Modal did not appear after clicking 'Class Execution' button.");
                if (!modalAppeared) {
                    console.error("Skipping a routine slot because the modal failed to load.");
                    continue;
                }
                const modal = document.querySelector('#myResultbacknew');

                const subjectSelect = await waitForElement('#subject_name_no', modal);
                const subjects = [];
                if (subjectSelect) {
                    for (let i = 1; i < subjectSelect.options.length; i++) {
                        subjects.push(subjectSelect.options[i].text.trim());
                    }
                }
                if (subjects.length === 0) subjects.push('Subject Not Found');
                
                const uniqueDates = new Set();
                document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                
                const dateInput = await waitForElement('input.datepicker_with_range', modal);
                if (dateInput) dateInput.click();
                
                await scrapeCalendarInDirection(uniqueDates, 'prev');

                const closeButtonForReset = modal.querySelector('button.close_btn');
                const modalDisappearedForReset = await clickThenPoll(closeButtonForReset, () => !document.querySelector('#myResultbacknew'));
                if(!modalDisappearedForReset) continue;

                const modalReappeared = await clickThenPoll(button, () => document.querySelector('#myResultbacknew'));
                if(!modalReappeared) continue;
                
                const reopenedModal = document.querySelector('#myResultbacknew');
                document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                
                const dateInput2 = await waitForElement('input.datepicker_with_range', reopenedModal);
                if (dateInput2) dateInput2.click();

                await scrapeCalendarInDirection(uniqueDates, 'next');

                const allDates = Array.from(uniqueDates);
                subjects.forEach(subject => {
                    allDates.forEach(date => {
                        data.push({ "Days": day, "Column": columnNumber, "Subject": subject, "Date": date });
                    });
                });
                
                const finalCloseButton = document.querySelector('#myResultbacknew button.close_btn');
                await clickThenPoll(finalCloseButton, () => !document.querySelector('#myResultbacknew'));

            } catch (error) {
                console.error("An error occurred during scraping a routine slot:", error);
                const openModal = document.querySelector('#myResultbacknew');
                if (openModal) openModal.querySelector('button.close_btn')?.click();
            }
        }
    
        if (data.length > 0) {
            alert(`Scraping complete! Found ${data.length} records. The file download will now begin.`);
            chrome.runtime.sendMessage({ action: 'download_data', data: data });
        } else {
            alert("Scraping finished, but no data was found.");
        }
    };


    // ===================================================================================
    // SECTION 3: FORM FILLING LOGIC
    // ===================================================================================

    const normalizeDataKeys = (data) => data.map(row => Object.keys(row).reduce((acc, key) => { acc[key.trim()] = row[key]; return acc; }, {}));

    const fillFormsFromExcel = async (data) => {
        let processedData = normalizeDataKeys(data);
        const routineTable = document.querySelector('.table.table-bordered');
        if (!routineTable) { alert("Fatal Error: Could not find routine table."); return; }

        const rowsToProcess = processedData.filter(row => String(row['Submitted'] || '').trim().toLowerCase() === 'yes' && ['yes', 'no'].includes(String(row['Class Execution'] || '').trim().toLowerCase()));

        if (rowsToProcess.length === 0) {
            alert("No records to process were found. Ensure 'Submitted' column is 'Yes'.");
            return;
        }
        alert(`Found ${rowsToProcess.length} records marked for processing. Automation will now begin.`);

        for (const rowData of rowsToProcess) {
            console.group(`Processing Excel Row for Date: ${rowData.Date}`);
            try {
                const tableRow = [...routineTable.querySelectorAll('tbody > tr')].find(r => r.cells[0]?.textContent.trim().toLowerCase() === rowData.Days.toLowerCase());
                if (!tableRow) throw new Error(`Could not find table row for day: ${rowData.Days}`);

                const execButton = tableRow.cells[parseInt(rowData.Column, 10)]?.querySelector('a.class_execution_data');
                if (!execButton) throw new Error(`Could not find 'Class Execution' button`);

                const modalAppeared = await clickThenPoll(execButton, () => document.querySelector('#myResultbacknew'), "Modal did not appear after clicking 'Class Execution' button.");
                if (!modalAppeared) throw new Error("Modal failed to load.");
                const modal = document.querySelector('#myResultbacknew');

                const datePicker = await pollForCondition(() => modal.querySelector('input.datepicker_with_range'));
                if(datePicker) { // Check if element exists before setting value
                    modal.querySelector('input.datepicker_with_range').value = rowData.Date;
                }
                
                const classExecDropdown = await pollForCondition(() => modal.querySelector('#class_execution_yes_no'));
                const executionStatus = String(rowData['Class Execution']).trim().toLowerCase();

                if (executionStatus === 'yes') {
                    classExecDropdown.value = '1';
                    classExecDropdown.dispatchEvent(new Event('change', { bubbles: true }));
                    const topicFieldAppeared = await pollForCondition(() => modal.querySelector('#class_execution_yes')?.offsetParent !== null, 5000, "Topic field did not appear.");
                    if (!topicFieldAppeared) throw new Error("Topic field failed to appear after selecting 'Yes'.");
                    if (!selectOptionByText(modal.querySelector('#subject_name_no'), rowData.Subject)) throw new Error(`Subject "${rowData.Subject}" could not be selected.`);
                    modal.querySelector('#class_execution_yes').value = rowData['Topic (Yes)'] || '';
                    modal.querySelector('#registerd_class').value = rowData['Total Students'] || '';
                    modal.querySelector('#attended_class').value = rowData['Attended Students'] || '';
                } else if (executionStatus === 'no') {
                    classExecDropdown.value = '0';
                    classExecDropdown.dispatchEvent(new Event('change', { bubbles: true }));
                    const reasonFieldAppeared = await pollForCondition(() => modal.querySelector('#class_execution_no')?.offsetParent !== null, 5000, "Reason field did not appear.");
                    if (!reasonFieldAppeared) throw new Error("Reason field failed to appear after selecting 'No'.");
                    if (!selectOptionByText(modal.querySelector('#class_execution_no'), rowData['Reason (No)'])) throw new Error(`Reason "${rowData['Reason (No)']}" could not be selected.`);
                }
                
                console.log("Submitting form automatically...");
                const submitButton = modal.querySelector('input.confirm_result_revart_back');
                const modalDisappeared = await clickThenPoll(submitButton, () => !document.querySelector('#myResultbacknew'), "Modal did not disappear after clicking submit. Submission may have failed.");
                if (modalDisappeared) {
                    rowData.Submitted = 'Processed (Success)';
                } else {
                    throw new Error("Modal did not close, indicating a possible submission failure.");
                }
            } catch (error) {
                console.error('An error occurred processing row:', error);
                rowData.Submitted = `Processed (Error: ${error.message})`;
                const openModal = document.querySelector('#myResultbacknew');
                if (openModal) openModal.querySelector('button.close_btn')?.click();
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