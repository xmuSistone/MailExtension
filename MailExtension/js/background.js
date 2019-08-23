// 全局的邮箱地址URL
let globalSubject = '您的工资条请查收'
let globalUrl263 = undefined
let globalUrlQq = undefined
let globalUrl163 = undefined
let globalUrl126 = undefined

/**
 * 接收消息, 然后回传tab地址 & 主题设置
 */
chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
    if (request.cmd === "setSubject") {
        let newSubject = request.newSubject
        globalSubject = newSubject
        sendResponse('0', request, sender)

        // 主题持久化到缓存
        chrome.storage.local.set({globalSubject: newSubject});
    } else if (request.cmd === "getSubject") {
        sendResponse(`${globalSubject}`, request, sender)
    } else if (request.cmd.startsWith('prepare')) {
        refreshMailTabUrl()
        sendResponse('1', request, sender)
    } else if (request.cmd === "tabUrl-qq") {
        let response = {}
        response.tabUrl = globalUrlQq
        response.subject = globalSubject
        sendResponse(response, request, sender)
    } else if (request.cmd === "tabUrl-263") {
        sendResponse(`${globalUrl263}`, request, sender)
    } else if (request.cmd === "tabUrl-163") {
        let response = {}
        response.tabUrl = globalUrl163
        response.subject = globalSubject
        sendResponse(response, request, sender)
    } else if (request.cmd === "tabUrl-126") {
        let response = {}
        response.tabUrl = globalUrl126
        response.subject = globalSubject
        sendResponse(response, request, sender)
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
            } else if (tabUrl.startsWith("https://mail.qq.com/cgi-bin/frame_html") || tabUrl.startsWith("http://mail.qq.com/cgi-bin/frame_html")) {
                globalUrlQq = tabUrl
            } else if (tabUrl.startsWith("https://mail.163.com/js6/main.jsp") || tabUrl.startsWith("http://mail.163.com/js6/main.jsp")) {
                globalUrl163 = tabUrl
            } else if (tabUrl.startsWith("https://mail.126.com/js6/main.jsp") || tabUrl.startsWith("http://mail.126.com/js6/main.jsp")) {
                globalUrl126 = tabUrl
            }
        }
    })
}

/**
 * 一上来就从缓存中读取主题设置
 */
chrome.storage.local.get('globalSubject', function(result) {
    if (result && result.globalSubject) {
        globalSubject = result.globalSubject
    }
});

// 启动时就检查下邮箱的URL地址
refreshMailTabUrl()

console.log('xxxx 到此一游')
