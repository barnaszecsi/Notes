function createTabSafely() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth();

    var suffix = "";

    // Logic for the name suffix
    if (month === 11 || month === 0) {
        var displayYear = (month === 11) ? year + 1 : year;
        suffix = displayYear + "/1";
    }
    else if (month === 5 || month === 6) {
        suffix = year + "/nyár";
    }
    else if (month === 7 || month === 8) {
        suffix = year + "/2";
    }
    else {
        suffix = year + "/generic-" + (month + 1);
    }

    var sheetName = "Részvétel " + suffix;
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        var template = ss.getSheetByName('Reszvetel Minta');
        if (template) {
            // 1. Create the copy
            var newSheet = template.copyTo(ss).setName(sheetName);

            // 2. Move it to the first position (index 1)
            ss.setActiveSheet(newSheet);
            ss.moveActiveSheet(1);

            Logger.log('Sheet created and moved to first position: ' + sheetName);
        } else {
            Logger.log('Error: Template "Reszvetel Minta" not found!');
        }
    } else {
        Logger.log('Sheet "' + sheetName + '" already exists!');
    }
}
