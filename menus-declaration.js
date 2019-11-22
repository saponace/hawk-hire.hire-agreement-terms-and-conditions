//0. On Open: Custom menu to run the script
function onOpen() {
    var ui = DocumentApp.getUi();
    ui.createMenu('Custom scripts')
        .addItem('Send agreement form', 'showEmailSelectionPopup')
        .addToUi();
}
