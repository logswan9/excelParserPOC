export class Base {

    getDateFromExcel(excelVal) {
        if (excelVal != "") {
            var date = new Date(0, 0, excelVal - 1, 0, 0, 0);
            var day = date.getDate();
            var month = date.getMonth() + 1;
            var year = date.getFullYear();
            date = (month + "/" + day + "/" + year);
            if (isNaN(day) || isNaN(month) || isNaN(year)){
                excelVal = String(excelVal);
            } else {
                excelVal = String(date);
            }
        } else {
            excelVal = "";
        }
        return excelVal;
    }

}