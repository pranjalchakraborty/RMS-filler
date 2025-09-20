document.addEventListener('DOMContentLoaded', () => {
    const scrapeBtn = document.getElementById('scrape');
    const fillBtn = document.getElementById('fill-from-excel');
    const fileInput = document.getElementById('excel-file-input');
  
    scrapeBtn.addEventListener('click', () => {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (tabs[0] && tabs[0].id) {
          chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            files: ['content.js']
          }).then(() => {
            console.log("Successfully injected content script from popup.");
            chrome.tabs.sendMessage(tabs[0].id, { action: 'scrape' });
          }).catch(err => {
            console.error("Popup failed to inject content script:", err);
            alert("Could not connect to the page. Please ensure you are on the correct routine management page and reload the page.");
          });
        } else {
            console.error("Could not find active tab.");
        }
      });
    });
  
    fillBtn.addEventListener('click', () => fileInput.click());
  
    fileInput.addEventListener('change', (event) => {
      const file = event.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
          chrome.runtime.sendMessage({
            action: 'parse_excel_data',
            data: e.target.result 
          });
          window.close();
        };
        reader.readAsDataURL(file);
      }
    });
});