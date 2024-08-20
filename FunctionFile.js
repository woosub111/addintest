Office.initialize = function () {
}

// 정보 표시줄에 상태 메시지를 추가하는 도우미 함수입니다.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function onButtonClick(event) {
    // 호출할 URL
    var url = "https://example.com/api/endpoint";

    // URL 호출을 위한 XMLHttpRequest
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, true);

    // 요청 헤더 설정 (필요 시)
    xhr.setRequestHeader("Content-Type", "application/json");

    // 응답 처리
    xhr.onreadystatechange = function () {
        if (xhr.readyState == 4 && xhr.status == 200) {
            // 요청이 성공했을 때 수행할 작업
            console.log("Request successful");
        } else if (xhr.readyState == 4) {
            // 요청이 실패했을 때 수행할 작업
            console.error("Request failed");
        }
    };

    // 요청 전송
    xhr.send();

    // 작업 완료를 알리기 위해 Office.js의 event.completed() 호출
    event.completed();
}
