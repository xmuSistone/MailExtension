﻿/**
 * excel列名，最多支持52列
 */
const SHEET_COLOMN = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', "AE", 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];

/**
 * toast配置 顶部居中
 */
toastr.options = {
    "positionClass": "toast-top-center",
};

/**
 * 监听DOM加载完成
 */
document.addEventListener('DOMContentLoaded', function () {
    // 往页面中注入一个工资助手的按钮
    injectSalaryButton()
});

/**
 * 在邮件的顶部加上 工资条助手按钮
 */
function injectSalaryButton() {
    let head_r_part_s = document.getElementsByClassName("head_r_part")
    if (head_r_part_s && head_r_part_s.length > 0) {
        let head_r_part = head_r_part_s[0]
        let nav_element = head_r_part.firstElementChild
        if (nav_element) {
            let liElement = document.createElement("li")
            liElement.onclick = onSalaryBtnClick
            liElement.style.width = '62px'
            liElement.style.height = '62px'

            let imgUrl = chrome.extension.getURL('../img/icon_white.png')
            liElement.style.backgroundImage = `url(${imgUrl})`
            liElement.style.backgroundPosition = 'center'
            liElement.style.backgroundRepeat = 'no-repeat'
            liElement.innerHTML = "<a href='#' style='background: #ffffff00;' title='工资条助手'></a>"

            // <li>标签中插入工资助手的按钮
            nav_element.insertBefore(liElement, nav_element.children[1])

            // 注入模态对话框
            injectDialogModal()
        }
    }
}

/**
 * 注入对话框
 */
function injectDialogModal() {
    let dialogUrl = chrome.runtime.getURL('../html/dialog.html')
    let xhr = new XMLHttpRequest();
    xhr.open('GET', dialogUrl, true);
    xhr.onreadystatechange = function () {
        let dialogContent = xhr.response
        if (dialogContent) {
            let modalDiv = document.getElementById("modal-1")
            if (!modalDiv) {
                // 注入对话框
                let divElement = document.createElement("div")
                divElement.setAttribute("id", "modal-1")
                divElement.setAttribute("aria-hidden", "true")
                divElement.setAttribute("class", "modal micromodal-slide")
                divElement.innerHTML = dialogContent
                document.body.appendChild(divElement)

                // 模板下载链接
                let templateHref = document.getElementById("salary-template-href")
                let salaryTemplateUrl = chrome.runtime.getURL('../res/salary-template.xls')
                templateHref.setAttribute("href", salaryTemplateUrl)

                // 一键发送按钮点击事件
                let sendMailBtn = document.getElementById("salary-dialog-confirm-btn")
                sendMailBtn.onclick = onSendMailBtnClick

                // 对话框初始化
                MicroModal.init();
            }
        }
    };
    xhr.send();
}

/**
 * 一键发送按钮 点击事件处理
 */
function onSendMailBtnClick() {
    let fileInput = document.getElementById("salary-dialog-file-input");
    let filesList = fileInput.files;
    if (!filesList || filesList.length === 0) {
        toastr.clear();
        toastr.info("请添加excel文件");
        return
    }

    // 读取文件内容
    let reader = new FileReader();
    reader.onload = function (e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, {type: 'binary'});
        let sheetNames = workbook.SheetNames;

        if (sheetNames && sheetNames.length > 0) {
            // 只处理第一个sheet
            let sheetName = sheetNames[0];
            let sheet = workbook.Sheets[sheetName];
            if (sheet) {
                parseExcel(sheet);
            }
        }
    };
    reader.readAsBinaryString(filesList[0]);

    MicroModal.close()
}

/**
 * 解析excel，默认最后一列是email地址
 */
function parseExcel(sheet) {
    // 1. 先提取出title
    let titleList = [];
    let maxColumn = SHEET_COLOMN.length;
    for (let i = 0; i < maxColumn; i++) {
        let columnsFirstKey = `${SHEET_COLOMN[i]}1`;
        let columnsFirstValue = sheet[columnsFirstKey];
        if (columnsFirstValue) {
            titleList.push(columnsFirstValue)
        } else {
            break;
        }
    }

    // 2. 所有要发的信息，存储到taskList
    let taskList = [];
    let columns = titleList.length;

    // 第一行是title，从第2行开始存储taskList
    let pointer = 2;
    while (pointer > 0) {
        let firstKey = `${SHEET_COLOMN[0]}${pointer}`;
        let firstValue = sheet[firstKey];
        if (firstValue) {
            // 本行第1列有值，需要把本行数据添加到taskList
            let itemTask = [];
            for (let column = 0; column < columns; column++) {
                let cellKey = `${SHEET_COLOMN[column]}${pointer}`;
                let cellValue = sheet[cellKey];
                itemTask.push(cellValue);
            }
            taskList.push(itemTask);
            pointer++;
        } else {
            break
        }
    }

    if (titleList.length > 0 && taskList.length > 0) {
        // 获取地址栏参数usr和sid
        chrome.runtime.sendMessage({cmd: 'tabUrl-263'}, function (response) {
            let tabUrlAddress = response.tabUrl;
            let usr = tabUrlAddress.match(/(usr=[0-9a-zA-Z]+@mistinechina.com)/g);
            let sid = tabUrlAddress.match(/(sid=[0-9a-zA-Z]+)/g);

            if (usr && usr.length > 0 && sid && sid.length > 0) {
                // 批量发送邮件
                globalTaskList = taskList
                globalTitleList = titleList
                globalSubject = response.subject
                updateNameIndex(titleList)
                globalUsr = usr[0]
                globalSid = sid[0]

                doSendMail(0)
            }
        });
    }
}

let nameIndex = 0
let globalTitleList = []
let globalTaskList = []
let globalSubject = ''
let globalUsr = ''
let globalSid = ''

/**
 * title中包含名字的index
 */
function updateNameIndex(titleList) {
    nameIndex = 0;
    let len = titleList.length;
    for (let i = 0; i < len; i++) {
        let titleCell = titleList[i];
        if (titleCell.w.indexOf('名') >= 0) {
            nameIndex = i;
            break;
        }
    }
}


/**
 * 点击工资条助手图标
 */
function onSalaryBtnClick() {
    // 即将弹出对话框，让background.js准备一些数据
    chrome.runtime.sendMessage({cmd: 'prepare-263'});
    MicroModal.show('modal-1');
}

/**
 * 真正发送email
 * @index 第几封邮件
 */
function doSendMail(index) {

    if (index >= globalTaskList.length) {
        return
    }
    let taskItem = globalTaskList[index]
    let emailAddr = taskItem[globalTitleList.length - 1]
    let mailContent = buildMailContent(globalTitleList, taskItem)


    // 1. 构造URL地址
    let timestamp = new Date().getTime()
    let fromMail = globalUsr.substring(4)
    let info = `${globalUsr}&${globalSid}&_count=1&_t=${timestamp}`
    let postUrl = `https://mail.263.net/wm2e/mail/mailOperate/mailOperateAction_mailSaveToFolder.do?${info}`

    // 2. 发送ajax请求
    let ajax = new XMLHttpRequest();
    ajax.onload = function () {
        let sendSuccess = false

        let response = ajax.response
        if (response) {
            let responseObj = JSON.parse(response)
            if (responseObj.success === true) {
                sendSuccess = true
            }
        }

        toastr.clear()
        if (sendSuccess) {
            // 邮件发送成功
            toastr.success(`第${index + 1}/${globalTaskList.length}封邮件发送成功 - ${taskItem[nameIndex].w}(${emailAddr.w})`)
            if (index < globalTaskList.length - 1) {
                doSendMail(index + 1)
            }
        } else {
            // 邮件发送失败
            toastr.error(`第${index + 1}/${globalTaskList.length}封邮件发送失败 - ${taskItem[nameIndex].w}(${emailAddr.w})，已中断`)
        }
    }

    ajax.open('post', postUrl);
    //post方式传递数据是模仿form表单传递给服务器的,要设置header头协议

    //3.发送请求(get--null    post--数据)
    let data = new FormData();
    // 3.1 谁发的？
    data.append('from', fromMail);
    // 3.2 发给谁
    data.append('tos', `${taskItem[nameIndex].w}<${emailAddr.w}>`);

    // 3.3 主题
    data.append('subject', globalSubject);

    // 3.4 邮件内容Html
    data.append('textInner', mailContent);

    // 3.6 其它无关紧要的一些参数，抓包见到的
    data.append('sendIndividually', 'false');
    data.append('replyPreviousEmailId', '');
    data.append('replyMailOpertType', '');
    data.append('mailOperateType', 'write');
    data.append('messageId', '');
    data.append('references', '');
    data.append('saveSend', 'on');
    data.append('emailpriority', '3');

    // 4. 真正发送
    ajax.send(data);
}

/**
 * 组装邮件正文
 * @param titleList 标题
 * @param taskItem 内容
 */
function buildMailContent(titleList, rowCells) {
    let count = titleList.length

    let mailContentStr = `<body><table style="width: 100%; background-color: white">\n`

    for (let i = 0; i < count; i = i + 2) {
        // 颜色重的div
        let titleCell_1 = titleList[i]
        let itemCell_1 = rowCells[i]
        let rowTr_1 =
            `<tr style="background-color: #eeeeee">
                <td style="font-size: 16px;padding: 10px;">${titleCell_1.w}</td>
                <td style="font-size: 16px;padding: 10px;">${itemCell_1.w}</td>
            </tr>\n`

        mailContentStr += rowTr_1

        // 颜色浅
        let titleCell_2 = titleList[i + 1]
        if (titleCell_2) {
            let itemCell_2 = rowCells[i + 1]
            let rowTr_2 =
                `<tr style="background-color: #f8f8f8">
                <td style="font-size: 16px;padding: 10px;">${titleCell_2.w}</td>
                <td style="font-size: 16px;padding: 10px;">${itemCell_2.w}</td>
            </tr>\n`
            mailContentStr += rowTr_2
        }
    }
    mailContentStr += "</table></body>"
    return mailContentStr
}
