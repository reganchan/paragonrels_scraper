let scrapeBtn = document.getElementById('scrapeBtn');

scrapeBtn.onclick = function(element) {
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
        chrome.tabs.executeScript(tabs[0].id, {file: "xlsx.full.min.js"});
        chrome.tabs.executeScript(tabs[0].id, {file: "scrape.js"});
    });
};