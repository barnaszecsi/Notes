function createTabSafely() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth();

    var datePart = "";
    var suffix = "";

    // 1. Determine Suffix and Logic
    if (month === 11 || month === 0) {
        var displayYear = (month === 11) ? year + 1 : year;
        datePart = displayYear + "/1";
        suffix = "/1";
    } else if (month === 5 || month === 6) {
        datePart = year + "/nyár";
        suffix = "/nyár";
    } else if (month === 7 || month === 8) {
        datePart = year + "/2";
        suffix = "/2";
    } else {
        datePart = year + "/generic-" + (month + 1);
    }

    var sheetName = "Részvétel " + datePart;
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        var template = ss.getSheetByName('Reszvetel Minta');
        if (template) {
            var newSheet = ss.insertSheet(sheetName, 0, { template: template });
            ss.setActiveSheet(newSheet);

            // 2. Fill Dates of Sundays based on specific suffixes
            fillSundays(newSheet, suffix, year);
        } else {
            Logger.log("Template sheet 'Reszvetel Minta' not found.");
        }
    }
    else {
        Logger.log("Sheet '" + sheetName + "' already exists.");
    }
}

function fillSundays(sheet, suffix, currentYear) {
    var startDate, endDate;

    // Ensure we use the correct year if it was incremented for the "/1" suffix
    if (suffix === "/1") {
        startDate = new Date(currentYear, 0, 1);
        endDate = new Date(currentYear, 5, 30);
    } else if (suffix === "/2") {
        startDate = new Date(currentYear, 7, 21);
        endDate = new Date(currentYear, 11, 24);
    } else if (suffix === "/nyár") {
        startDate = new Date(currentYear, 6, 1);
        endDate = new Date(currentYear, 7, 19);
    } else { return; }

    var sundays = [];
    var d = new Date(startDate);
    while (d <= endDate) {
        if (d.getDay() === 0) {
            sundays.push(new Date(d));
        }
        d.setDate(d.getDate() + 1);
    }

    if (sundays.length > 0) {
        var neededCols = 4 + sundays.length;
        var currentCols = sheet.getMaxColumns();
        if (neededCols > currentCols) {
            sheet.insertColumnsAfter(currentCols, neededCols - currentCols);
        }

        var calendars = CalendarApp.getCalendarsByName('Holidays in Hungary');
        var holidayCal = (calendars.length > 0) ? calendars[0] : null;
        if (!holidayCal) {
            holidayCal = CalendarApp.getCalendarById('hu.hungarian#holiday@group.v.calendar.google.com');
        }

        // Clear row 17 from column D onwards
        var lastCol = sheet.getMaxColumns();
        if (lastCol >= 4) {
            sheet.getRange(17, 4, 1, lastCol - 3).clearContent().setBackground(null).setFontColor(null);
        }

        for (var i = 0; i < sundays.length; i++) {
            var col = 4 + i;
            var dateCell = sheet.getRange(1, col);
            var infoCell = sheet.getRange(17, col);
            var sundayDate = sundays[i];

            dateCell.setValue(sundayDate).setNumberFormat("yyyy.MM.dd");

            var holidayNames = [];
            if (holidayCal) {
                for (var offset = -3; offset <= 2; offset++) {
                    var checkDate = new Date(sundayDate);
                    checkDate.setDate(sundayDate.getDate() + offset);

                    var events = holidayCal.getEventsForDay(checkDate);
                    if (events.length > 0) {
                        events.forEach(function (e) {
                            var name = e.getTitle();
                            var isExcluded = name.includes("Nicholas") ||
                                name.includes("Father’s Day") ||
                                name.includes("Miklós") ||
                                name.includes("Apák napja");

                            if (!isExcluded && holidayNames.indexOf(name) === -1) {
                                holidayNames.push(name);
                            }
                        });
                    }
                }
            }

            // Fill logic for Row 17
            if (holidayNames.length > 0) {
                // HOLIDAY FOUND
                dateCell.setBackground("#ff0000");
                infoCell.setValue(holidayNames.join(", "))
                    .setFontColor("#ff0000")
                    .setWrap(true);
            } else {
                // NO HOLIDAY -> SCH
                infoCell.setValue("SCH")
                    .setFontColor("#000000")
                    .setFontWeight("normal");
            }
        }
    }
}