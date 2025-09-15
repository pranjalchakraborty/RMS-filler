(function() {
    'use strict';

    // ===================================================================================
    // SECTION 1: CORE UTILITY FUNCTIONS (Unchanged)
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

    const clickAndWait = async (element, delay = 1500) => {
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
    // SECTION 2: DATA SCRAPING LOGIC (With Upgraded Calendar and Subject Scraping)
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

    /**
     * NEW: Upgraded function to scrape both past and future dates without duplicates.
     * @param {Set<string>} dateSet - A Set object to store unique dates.
     * @param {'prev' | 'next'} direction - The direction to navigate ('prev' for past, 'next' for future).
     */
    const scrapeCalendarInDirection = async (dateSet, direction) => {
        let calendar = await waitForElement('.datepicker.datepicker-dropdown');
        if (!calendar) return; // Exit if calendar not found

        while (true) {
            // Scrape dates from the currently visible month
            calendar.querySelectorAll('.day.allowed-date:not(.old):not(.new)').forEach(el => {
                const timestamp = el.getAttribute('data-date');
                if (timestamp) {
                    const date = new Date(parseInt(timestamp));
                    // Format date as YYYY-MM-DD and add to the Set
                    dateSet.add(date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2));
                }
            });

            // Find the navigation button for the specified direction
            const navButton = calendar.querySelector(`.${direction}:not(.disabled)`);
            if (navButton) {
                await clickAndWait(navButton, 500); // Click and wait for the calendar to update
                calendar = await waitForElement('.datepicker.datepicker-dropdown'); // Re-find the calendar
                if (!calendar) break;
            } else {
                break; // Exit loop if the button is disabled or not found
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

                await clickAndWait(button);
                const modal = await waitForElement('#myResultbacknew.in');
                if (!modal) continue;

                // --- CHANGE 2: Scrape ALL subjects from the dropdown ---
                const subjects = [];
                const subjectSelect = await waitForElement('#subject_name_no', modal);
                if (subjectSelect) {
                    // Start from index 1 to skip the "--Select Subject--" placeholder
                    for (let i = 1; i < subjectSelect.options.length; i++) {
                        subjects.push(subjectSelect.options[i].text.trim());
                    }
                }
                if (subjects.length === 0) subjects.push('Subject Not Found'); // Fallback
                
                // --- CHANGE 3: Scrape both past and future dates ---
                const uniqueDates = new Set();
                
                // Phase 1: Scrape backwards (past)
                document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                await showCalendarProgrammatically();
                await scrapeCalendarInDirection(uniqueDates, 'prev');

                // Phase 2: Reset and scrape forwards (future)
                // We close and reopen the modal to reset the calendar to the default month
                await clickAndWait(modal.querySelector('button.close_btn'), 500);
                await clickAndWait(button); // Reopen the same modal
                const reopenedModal = await waitForElement('#myResultbacknew.in');
                if (!reopenedModal) continue;
                
                document.querySelectorAll('.datepicker.datepicker-dropdown').forEach(cal => cal.remove());
                await showCalendarProgrammatically();
                await scrapeCalendarInDirection(uniqueDates, 'next');

                const allDates = Array.from(uniqueDates);

                // --- Combine all subjects with all dates ---
                subjects.forEach(subject => {
                    allDates.forEach(date => {
                        data.push({ "Days": day, "Column": columnNumber, "Subject": subject, "Date": date });
                    });
                });

                await clickAndWait(reopenedModal.querySelector('button.close_btn'), 500);
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
    // SECTION 3: FORM FILLING LOGIC (With "Yes" Trigger)
    // ===================================================================================
    
    const normalizeDataKeys = (data) => data.map(row => Object.keys(row).reduce((acc, key) => { acc[key.trim()] = row[key]; return acc; }, {}));

    const fillFormsFromExcel = async (data) => {
        let processedData = normalizeDataKeys(data);
        const routineTable = document.querySelector('.table.table-bordered');
        if (!routineTable) { alert("Fatal Error: Could not find routine table."); return; }

        // --- CHANGE 1: Filter logic now checks for "yes" in the 'Submitted' column ---
        const rowsToProcess = processedData.filter(row => {
            const submittedStatus = String(row['Submitted'] || '').trim().toLowerCase();
            const executionStatus = String(row['Class Execution'] || '').trim().toLowerCase();
            return submittedStatus === 'yes' && ['yes', 'no'].includes(executionStatus);
        });

        if (rowsToProcess.length === 0) { 
            alert("No records to process were found. Ensure the 'Submitted' column is marked as 'Yes' for the rows you wish to process."); 
            return; 
        }
        
        alert(`Found ${rowsToProcess.length} records marked for processing. The automation will now begin.`);

        for (const rowData of rowsToProcess) {
            console.group(`Processing Excel Row for Date: ${rowData.Date}`);
            try {
                const tableRow = [...routineTable.querySelectorAll('tbody > tr')].find(r => r.cells[0]?.textContent.trim().toLowerCase() === rowData.Days.toLowerCase());
                if (!tableRow) throw new Error(`Could not find table row for day: ${rowData.Days}`);

                const execButton = tableRow.cells[parseInt(rowData.Column, 10)]?.querySelector('a.class_execution_data');
                if (!execButton) throw new Error(`Could not find 'Class Execution' button`);

                await clickAndWait(execButton, 2000);
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
                
                // Auto-submit without confirmation
                console.log("Submitting form automatically...");
                await clickAndWait(modal.querySelector('input.confirm_result_revart_back'), 2500);
                rowData.Submitted = 'Processed (Success)'; // Update status for the final report
                
            } catch (error) {
                console.error('An error occurred processing row:', error);
                rowData.Submitted = `Processed (Error: ${error.message})`; // Update status with the error
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
    // SECTION 4: MAIN MESSAGE LISTENER (Unchanged)
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