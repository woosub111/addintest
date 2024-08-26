Office.initialize = function () {
}


/**
 * Writes 'Hello world!' to a new message body.
 */
function insertTemplate(event) {
    const template = "<p>안녕하세요,</p><p>이것은 템플릿 내용입니다.</p>";
    Office.context.mailbox.item.body.setAsync(
        template,
        {
            coercionType: Office.CoercionType.Html
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("템플릿이 성공적으로 삽입되었습니다.");
            } else {
                console.error("템플릿 삽입 중 오류 발생:", asyncResult.error);
            }
            event.completed();
        }
    );
}

function openWebsite(event) {
    window.open("https://hsi.cleverse.kr/externalHome", "_blank");
}

function reportSpam(event) {
    const itemId = Office.context.mailbox.item.itemId;
    callApi('/reportSpam', { itemId: itemId })
        .then(() => moveToJunkFolder())
        .then(() => {
            console.log("스팸 신고 및 이동 완료");
            event.completed();
        })
        .catch(error => {
            console.error("스팸 처리 중 오류 발생:", error);
            event.completed();
        });
}


function callApi(endpoint, data) {
    return fetch(`https://api.yourdomain.com${endpoint}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
    }).then(response => response.json());
},

function moveToJunkFolder() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.moveAsync("junkemail", function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            } else {
                reject(asyncResult.error);
            }
        });
    });
}


/*const MyAddIn = {
    openWebsite: function (event) {
        const url = "https://www.naver.com";
        window.open(url, '_blank');
        event.completed();
    },
    
    insertTemplate: function (event) {
        const template = "<p>안녕하세요,</p><p>이것은 템플릿 내용입니다.</p>";
        Office.context.mailbox.item.body.setAsync(
            template,
            { coercionType: Office.CoercionType.Html },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("템플릿이 성공적으로 삽입되었습니다.");
                } else {
                    console.error("템플릿 삽입 중 오류 발생:", asyncResult.error);
                }
                event.completed();
            }
        );
    },

    reportSpam: function (event) {
        const itemId = Office.context.mailbox.item.itemId;
        MyAddIn.callApi('/reportSpam', { itemId: itemId })
            .then(() => MyAddIn.moveToJunkFolder())
            .then(() => {
                console.log("스팸 신고 및 이동 완료");
                event.completed();
            })
            .catch(error => {
                console.error("스팸 처리 중 오류 발생:", error);
                event.completed();
            });
    },

    callApi: function (endpoint, data) {
        return fetch(`https://api.yourdomain.com${endpoint}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        }).then(response => response.json());
    },

    moveToJunkFolder: function () {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.moveAsync("junkemail", function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(asyncResult.error);
                }
            });
        });
    }
};

function insertTemplate(event) {
   Office.context.mailbox.item.body.setAsync(
        "Hello world!",
        {
            coercionType: "html", // Write text as HTML
        },

        // Callback method to check that setAsync succeeded
        function (asyncResult) {
            if (asyncResult.status ==
                Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
        }
    );
    /*
    const template = "<p>안녕하세요,</p><p>이것은 템플릿 내용입니다.</p>";
    Office.context.mailbox.item.body.setAsync(
        template,
        {
            coercionType: Office.CoercionType.Html
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("템플릿이 성공적으로 삽입되었습니다.");
            } else {
                console.error("템플릿 삽입 중 오류 발생:", asyncResult.error);
            }
            event.completed();
        }
    );
    
}

Office.onReady(() => {
    Office.actions.associate("openWebsite", MyAddIn.openWebsite);
    Office.actions.associate("insertTemplate", insertTemplate);
    Office.actions.associate("reportSpam", MyAddIn.reportSpam);
});*/
