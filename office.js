Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";

    }
});
Office.initialize = function () {
    window.alert = function (message) {
        app.showNotification("Title For the Notification", message)
    };
};