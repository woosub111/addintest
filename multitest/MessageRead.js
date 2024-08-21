(function () {
  "use strict";

  var messageBanner;

  // 새 페이지가 로드될 때마다 Office 초기화 함수가 실행되어야 합니다.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
  };

  // AttachmentDetails 개체의 배열을 가져와 줄 바꿈으로 구분된 첨부 파일 이름 목록을 작성합니다.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // EmailAddressDetails 개체를 다음으로 형식 지정
  // 이름 성 <이메일 주소>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // EmailAddressDetails 개체의 배열을 가져와
  // 줄 바꿈으로 구분된 형식 문자열 목록을 작성합니다.
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

    function reportSpam(event) {
        const itemId = Office.context.mailbox.item.itemId;
        // 스팸 신고를 처리하는 API 호출 예시
        fetch('https://api.yourdomain.com/reportSpam', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ itemId: itemId })
        })
            .then(response => response.json())
            .then(data => {
                console.log("스팸 신고 완료:", data);
                Office.context.mailbox.item.moveAsync("junkemail", function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("이메일이 스팸 폴더로 이동되었습니다.");
                    } else {
                        console.error("이메일 이동 중 오류 발생:", asyncResult.error);
                    }
                    event.completed();
                });
            })
            .catch(error => {
                console.error("스팸 신고 중 오류 발생:", error);
                event.completed();
            });
    }


  // 항목 기반 개체에서 속성을 로드한 다음
  // 메시지 관련 속성을 로드합니다.
  function loadProps() {
    var item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // 알림 표시를 위한 도우미 함수입니다.
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();