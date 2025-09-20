// ===================================================================================
// THE SINGLETON GUARD: Ensures this script's logic only runs once per page.
// ===================================================================================
if (window.myRoutineScraperHasBeenInjected) {
    console.log("DEBUG: content.js was already injected. Halting to prevent duplicates.");
} else {
    window.myRoutineScraperHasBeenInjected = true;
    console.log("DEBUG: content.js injected and initialized for the first time.");

    (function() {
        'use strict';

        // ===================================================================================
        // SECTION 0: THE GATEKEEPER FLAG
        // ===================================================================================
        let isProcessRunning = false;

        // ===================================================================================
        // SECTION 1: CORE UTILITY FUNCTIONS
        // ===================================================================================

        const waitForElement = (selector, context = document, timeout = 30000) => {
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
        
        const clickAndWait = async (element, delay = 1500) => {
            if (element) {
                element.click();
                await new Promise(resolve => setTimeout(resolve, delay));
            }
        };
        
        const safeClick = (elementToClick) => {
            return new Promise(resolve => {
                if (!elementToClick) {
                    console.error("safeClick failed: element is null.");
                    return resolve();
                }
                const clickHandler = (event) => {
                    event.preventDefault();
                    elementToClick.removeEventListener('click', clickHandler, true);
                };
                elementToClick.addEventListener('click', clickHandler, true);
                const event = new MouseEvent('click', { bubbles: true, cancelable: true, view: window });
                elementToClick.dispatchEvent(event);
                setTimeout(resolve, 100);
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

        // ===================================================================================
        // SECTION 2: DATA SCRAPING LOGIC
        // ===================================================================================
        
        const openCalendar = async (modal) => {
            const dateInput = await waitForElement('input.datepicker_with_range', modal);
            if (dateInput) {
                console.log("DEBUG: Found date input. Focusing and clicking to open calendar.");
                dateInput.focus();
                dateInput.click();
                await new Promise(r => setTimeout(r, 300));
            } else {
                console.error("DEBUG: Could not find date input to open calendar.");
            }
        };

        const scrapeCalendarInDirection = async (dateSet, direction) => {
            let calendar = await waitForElement('.datepicker.datepicker-dropdown');
            if (!calendar) {
                console.error("DEBUG: Calendar did not appear for scraping.");
                return;
            }
            console.log(`DEBUG: Starting calendar scrape in direction: ${direction}`);
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
                    await clickAndWait(navButton, 500);
                    calendar = await waitForElement('.datepicker.datepicker-dropdown');
                    if (!calendar) break;
                } else {
                    console.log(`DEBUG: No more '${direction}' button. Ending scrape in this direction.`);
                    break;
                }
            }
        };

        const scrapeRoutine = async () => {
            try {
                console.log("DEBUG: scrapeRoutine() called.");
                alert("Starting to scrape routine data. Please do not interact with the page until the final report is downloaded.");
                const data = [];
                
                console.log("DEBUG: Finding all 'a.class_execution_data' buttons.");
                const buttons = [...document.querySelectorAll('a.class_execution_data')];
                console.log(`DEBUG: Found ${buttons.length} buttons to process.`);
            
                let buttonCounter = 0;
                for (const button of buttons) {
                    buttonCounter++;
                    const buttonRelId = button.getAttribute('rel');
                    console.log(`%cDEBUG: Starting loop iteration ${buttonCounter} for button with rel="${buttonRelId}".`, 'font-weight: bold; color: blue;');

                    try {
                        const cell = button.closest('td');
                        const row = button.closest('tr');
                        if (!row || !cell) continue;
                        const columnNumber = cell.cellIndex;
                        const day = row.cells[0]?.textContent.trim();
        
                        await safeClick(button);
                        console.log(`DEBUG: safeClick executed for button ${buttonRelId}. Polling for modal...`);
                        const modal = await waitForElement('#myResultbacknew.in');
                        if (!modal) {
                            console.error(`DEBUG: Modal did not appear for button ${buttonRelId}. Skipping.`);
                            continue;
                        }
                        console.log(`DEBUG: Modal appeared for button ${buttonRelId}.`);
        
                        const subjects = [];
                        const subjectSelect = await waitForElement('#subject_name_no', modal);
                        if (subjectSelect) {
                            for (let i = 1; i < subjectSelect.options.length; i++) {
                                subjects.push(subjectSelect.options[i].text.trim());
                            }
                        }
                        if (subjects.length === 0) subjects.push('Subject Not Found');
                        
                        const uniqueDates = new Set();
                        
                        console.log("DEBUG: Opening calendar for the first time (past dates).");
                        document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                        await openCalendar(modal);
                        await scrapeCalendarInDirection(uniqueDates, 'prev');
        
                        const closeButtonForReset = modal.querySelector('button.close_btn');
                        console.log("DEBUG: Closing modal to reset for future date scraping.");
                        await clickAndWait(closeButtonForReset, 1000);
                        
                        console.log("DEBUG: Re-finding button to avoid stale element reference.");
                        const freshButton = document.querySelector(`a.class_execution_data[rel="${buttonRelId}"]`);
                        if (!freshButton) {
                            console.error(`DEBUG: CRITICAL FAILURE - Could not re-find button with rel="${buttonRelId}". Loop will likely fail. Skipping.`);
                            continue;
                        }
                        
                        console.log("DEBUG: Re-opening modal for future date scraping.");
                        await safeClick(freshButton);
                        const reopenedModal = await waitForElement('#myResultbacknew.in');
                        if (!reopenedModal) {
                             console.error(`DEBUG: Modal did not RE-APPEAR for button ${buttonRelId}. Skipping.`);
                            continue;
                        }
                        console.log(`DEBUG: Modal RE-APPEARED for button ${buttonRelId}.`);
                        
                        document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                        await openCalendar(reopenedModal);
                        await scrapeCalendarInDirection(uniqueDates, 'next');
        
                        const allDates = Array.from(uniqueDates);
                        subjects.forEach(subject => {
                            allDates.forEach(date => {
                                data.push({ "Days": day, "Column": columnNumber, "Subject": subject, "Date": date });
                            });
                        });

                        console.log(`DEBUG: Closing modal finally for button ${buttonRelId}.`);
                        await clickAndWait(reopenedModal.querySelector('button.close_btn'), 500);

                    } catch (error) {
                        console.error(`DEBUG: An error occurred inside the loop for button ${buttonRelId}:`, error);
                        const openModal = document.querySelector('#myResultbacknew.in');
                        if (openModal) await clickAndWait(openModal.querySelector('button.close_btn'), 500);
                    }
                    console.log(`%cDEBUG: Finished loop iteration ${buttonCounter} for button ${buttonRelId}.`, 'font-weight: bold; color: green;');
                }
            
                console.log("DEBUG: The main scraping loop has finished.");
                if (data.length > 0) {
                    alert(`Scraping complete! Found ${data.length} records. The file download will now begin.`);
                    chrome.runtime.sendMessage({ action: 'download_data', data: data });
                } else {
                    alert("Scraping finished, but no data was found. Please check the console for errors.");
                }
            } finally {
                isProcessRunning = false;
                console.log("DEBUG: Scraping process finished. Gatekeeper unlocked.");
            }
        };
        
        // ===================================================================================
        // SECTION 3: FORM FILLING LOGIC (Hardened and Logged)
        // ===================================================================================
        
        const normalizeDataKeys = (data) => data.map(row => Object.keys(row).reduce((acc, key) => { acc[key.trim()] = row[key]; return acc; }, {}));

        const fillFormsFromExcel = async (data) => {
            try {
                console.log("DEBUG: fillFormsFromExcel() called.");
                let processedData = normalizeDataKeys(data);
                const routineTable = document.querySelector('.table.table-bordered');
                if (!routineTable) { alert("Fatal Error: Could not find routine table."); return; }
                const rowsToProcess = processedData.filter(row => {
                    const submittedStatus = String(row['Submitted'] || '').trim().toLowerCase();
                    const executionStatus = String(row['Class Execution'] || '').trim().toLowerCase();
                    return submittedStatus === 'yes' && ['yes', 'no'].includes(executionStatus);
                });
                if (rowsToProcess.length === 0) { 
                    alert("No records to process were found. Ensure the 'Submitted' column is marked as 'Yes'."); 
                    return; 
                }
                alert(`Found ${rowsToProcess.length} records marked for processing. The automation will now begin.`);
                for (const rowData of rowsToProcess) {
                    console.group(`DEBUG: Processing Excel Row for Date: ${rowData.Date}`);
                    try {
                        const tableRow = [...routineTable.querySelectorAll('tbody > tr')].find(r => r.cells[0]?.textContent.trim().toLowerCase() === rowData.Days.toLowerCase());
                        if (!tableRow) throw new Error(`Could not find table row for day: ${rowData.Days}`);
                        const execButton = tableRow.cells[parseInt(rowData.Column, 10)]?.querySelector('a.class_execution_data');
                        if (!execButton) throw new Error(`Could not find 'Class Execution' button`);
        
                        await safeClick(execButton);
                        const modal = await waitForElement('#myResultbacknew.in');
                        if (!modal) throw new Error("Execution modal did not open.");
        
                        (await waitForElement('input.datepicker_with_range', modal)).value = rowData.Date;
                        const classExecDropdown = await waitForElement('#class_execution_yes_no', modal);
                        const executionStatus = String(rowData['Class Execution']).trim().toLowerCase();
                        if (executionStatus === 'yes') {
                            classExecDropdown.value = '1';
                            classExecDropdown.dispatchEvent(new Event('change', { bubbles: true }));
                            await new Promise(r => setTimeout(r, 500)); 
                            if (!selectOptionByText(await waitForElement('#subject_name_no', modal), rowData.Subject)) throw new Error(`Subject "${rowData.Subject}" could not be selected.`);
                            (await waitForElement('#class_execution_yes', modal)).value = rowData['Topic (Yes)'] || '';
                            (await waitForElement('#registerd_class', modal)).value = rowData['Total Students'] || '';
                            (await waitForElement('#attended_class', modal)).value = rowData['Attended Students'] || '';
                        } else if (executionStatus === 'no') {
                            classExecDropdown.value = '0';
                            classExecDropdown.dispatchEvent(new Event('change', { bubbles: true }));
                            await new Promise(r => setTimeout(r, 500));
                            if (!selectOptionByText(await waitForElement('#class_execution_no', modal), rowData['Reason (No)'])) throw new Error(`Reason "${rowData['Reason (No)']}" could not be selected.`);
                        }
                        
                        console.log("DEBUG: Submitting form automatically...");
                        await clickAndWait(modal.querySelector('input.confirm_result_revart_back'), 2500);
                        rowData.Submitted = 'Processed (Success)';
                        
                    } catch (error) {
                        console.error('DEBUG: An error occurred processing row:', error);
                        rowData.Submitted = `Processed (Error: ${error.message})`;
                        const openModal = document.querySelector('#myResultbacknew.in');
                        if (openModal) await clickAndWait(openModal.querySelector('button.close_btn'), 500);
                    } finally {
                        console.groupEnd();
                    }
                }
                
                alert("Finished processing all entries. The updated Excel report will now be downloaded.");
                chrome.runtime.sendMessage({ action: 'download_updated_excel', data: processedData });
            } finally {
                isProcessRunning = false;
                console.log("DEBUG: Form-filling process finished. Gatekeeper unlocked.");
            }
        };

        // ===================================================================================
        // SECTION 4: MAIN MESSAGE LISTENER
        // ===================================================================================

        chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
            console.log(`DEBUG: Message received in content script:`, request);
            if (isProcessRunning) {
                console.warn("DEBUG: A process is already running. Ignoring new request.");
                alert("A process is already running. Please wait for it to complete before starting a new one.");
                sendResponse({ status: 'already running' });
                return true;
            }
            if (request.action === 'scrape' || request.action === 'fill_form_data') {
                isProcessRunning = true;
                console.log(`DEBUG: Gatekeeper locked. Starting process: ${request.action}`);
                if (request.action === 'scrape') {
                    scrapeRoutine();
                } else if (request.action === 'fill_form_data') {
                    fillFormsFromExcel(request.data);
                }
            }
            sendResponse({ status: 'action received' });
            return true; 
        });

    })();
}