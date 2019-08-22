

document.getElementById("subject_confirm_btn").onclick = () => {
    let subjectInput = document.getElementById("subject_input")
    let newSubject = subjectInput.value

    if (newSubject) {
        subjectInput.placeholder = newSubject
        subjectInput.value = ''
        chrome.runtime.sendMessage({cmd: 'setSubject', newSubject: newSubject});
    }
}


chrome.runtime.sendMessage({cmd: 'getSubject'}, function (response) {
    let subjectInput = document.getElementById("subject_input")
    subjectInput.placeholder = response
});