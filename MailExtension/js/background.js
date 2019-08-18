// 全局的邮箱地址URL
let globalUrl263 = undefined
let globalUrlQq = undefined

/**
 * 接收消息，获取Cookie, 然后回传
 */
chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
    if (request.cmd === 'prepare-263') {
        refreshMailTabUrl()
        sendResponse('1', request, sender)
    } else if (request.cmd === "tabUrl-263") {
        sendResponse(`${globalUrl263}`, request, sender)
    } if (request.cmd === 'prepare-qq') {
        refreshMailTabUrl()
        sendResponse('1', request, sender)
    } else if (request.cmd === "tabUrl-qq") {
        sendResponse(`${globalUrlQq}`, request, sender)
    } else {
        sendResponse('0', request, sender);
    }
});

function refreshMailTabUrl() {
    chrome.tabs.query({}, function (allTabs) {
        for (let itemTab of allTabs) {
            let tabUrl = itemTab.url
            if (tabUrl.startsWith("https://mail.263.net/wm2e/mail/login/show/loginShowAction_loginShow.do")) {
                globalUrl263 = tabUrl
            } else if (tabUrl.startsWith("https://mail.qq.com/") || tabUrl.startsWith("http://mail.qq.com/")) {
                globalUrlQq = tabUrl
            }
        }
    })
}

// 启动时就检查下邮箱的URL地址
refreshMailTabUrl()
