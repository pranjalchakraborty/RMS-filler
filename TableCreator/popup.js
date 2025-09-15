document.addEventListener('DOMContentLoaded', () => {
  const scrapeBtn = document.getElementById('scrape');
  const fillBtn = document.getElementById('fill-from-excel');
  const fileInput = document.getElementById('excel-file-input');

  scrapeBtn.addEventListener('click', () => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (tabs[0]) chrome.tabs.sendMessage(tabs[0].id, { action: 'scrape' });
    });
  });

  fillBtn.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        // e.target.result will now be a base64 data URL string
        chrome.runtime.sendMessage({
          action: 'parse_excel_data',
          data: e.target.result 
        });
        window.close();
      };

      // *** CHANGE HERE: Read the file as a Data URL (base64) instead of an ArrayBuffer ***
      reader.readAsDataURL(file);
    }
  });
});