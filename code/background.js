chrome.runtime.onInstalled.addListener(() => {
	chrome.contextMenus.create({
		id: 'gn:simple-export-table',
		title: 'Export table to excel',
		contexts: ['all']
	});
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
    if (info.menuItemId.endsWith('gn:simple-export-table')) {
        const completionData = await chrome.tabs.executeScript(tab.id, {
            frameId: info.frameId,
            file: "/read-content.js"
        });
    }
});