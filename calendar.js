/*
    Adds a custom Google Docs menu to generate calendars,
    and automatically highlight the current day.
*/

function onInstall() {
    onOpen();
}
  
function daysInMonth(month, year) {
    return new Date(year, month, 0).getDate();
}

function generateCalendar(){
    var monthBegin = new Date()
    var dayOfMonth = new Date().getDate();
    monthBegin.setDate(monthBegin.getDate() - (dayOfMonth - 1));
    var monthBeginDay = new Date(monthBegin).getDay();
    x = 0;
    calendar = [];
    row = [];

    while(x < monthBeginDay){
        row.push('');
        x++;
    }

    x = 1;

    while(row.length < 7){
        row.push(x);
        x++;
    }

    calendar.push(row)

    var currentMonth = new Date().getMonth();
    var daysInMonthVal = daysInMonth(currentMonth, 2018);

    y = 0;

    while(x < daysInMonthVal){

        newRow = [];
        while(newRow.length < 7 && x <= daysInMonthVal){
        newRow.push(x);
        x++;
        }
        calendar.push(newRow);
    }
    return calendar
}

function getRowByDay(monthBeginDay){
    calendar = generateCalendar();
    var dayOfMonth = new Date().getDate();
    var dayRow = null;

    for(var z = 0; z < calendar.length; z++){
        if(calendar[z].indexOf(dayOfMonth) > -1){
        dayRow = z;
        }
    }

    return dayRow;
}

function addCalendar(){
    calendar = generateCalendar();
    var monthNames = ["January", "February", "March", "April",
        "May", "June","July", "August", "September", "October",
        "November", "December"];
    var monthBegin = new Date()
    var thisMonth = monthNames[monthBegin.getMonth()]
    var body = DocumentApp.getActiveDocument().getBody();
    body.appendHorizontalRule();
    var header = body.appendParagraph(thisMonth);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING4);
    var calendarTable = body.appendTable(calendar)
    var tableStyle = {};
    tableStyle[DocumentApp.Attribute.BOLD] = true;  
    tableStyle[DocumentApp.Attribute.MINIMUM_HEIGHT]= 3;
    calendarTable.setAttributes(tableStyle);
}

function updateCalendar(calendar){
    var monthBegin = new Date()
    var dayOfMonth = new Date().getDate();
    var dayOfWeek = new Date().getDay();
    monthBegin.setDate(monthBegin.getDate() - (dayOfMonth - 1));
    var monthBeginDay = new Date(monthBegin).getDay();
    var numRows = calendar.getNumRows();
    var dayRow = getRowByDay(monthBeginDay);
    var cell = calendar.getCell(dayRow, dayOfWeek);

    if(dayOfWeek == 0){
        var cellBefore = calendar.getCell(dayRow-1, 6);
    }else{
        var cellBefore = calendar.getCell(dayRow, dayOfWeek - 1);
    }

    var cellStyle = {};
    cellStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FF6666";  
    cell.setAttributes(cellStyle);

    var cellBeforeStyle = {};
    cellBeforeStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FFFFFF";  
    cellBefore.setAttributes(cellBeforeStyle);
}


function onOpen(){
    var body = DocumentApp.getActiveDocument().getBody();
    var table = body.getTables()[0];
    updateCalendar(table);
    var ui = DocumentApp.getUi();
    ui.createMenu('Meghan\'s Menu')
        .addItem('Generate This Month\'s Calendar!', 'addCalendar')
        .addToUi();
}
