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
    // 새 탭에서 열 URL
    var url = "https://hsi.cleverse.kr/externalHome";

    // 새 브라우저 탭에서 URL 열기
    window.open(url, '_blank');

    // 작업 완료를 알리기 위해 Office.js의 event.completed() 호출
    event.completed();
}
