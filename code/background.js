// Add context menu
chrome.contextMenus.create({
    id: 'export-table',
    title: 'Export table to excel',
    contexts: ['all']
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
    if (info.menuItemId.endsWith('export-table')) {
        const completionData = await chrome.tabs.executeScript(tab.id, {
            frameId: info.frameId,
            file: "/read-content.js"
        });
    }
});