
function main(workbook: ExcelScript.Workbook) {
    let serviceSheet = workbook.getWorksheet("ServiceSheet");
    let rosterSheet = workbook.getWorksheet("RosterSheet");
    let totalSheet = workbook.getWorksheet("TotalSheet");

    let serviceTable = serviceSheet.getTable("Service");
    let rosterTable = rosterSheet.getTable("Roster");

    // dictionary in form of {email: total service hours}
    let totalService: Map<string, number> = new Map();

    let serviceVals = serviceTable.getRange().getValues();

    // Grab all emails
    for (let n = 1; n < serviceVals.length; n++) {
        let email = serviceVals[n][3].toString();
        totalService.set(email, 0 as number);
    }

    // Add the service hours up for the people
    for (let n = 1; n < serviceVals.length; n++) {
        let email = serviceVals[n][3].toString();
        let serviceAmt = Number(serviceVals[n][6]); // I hate javascript
        totalService.set(email, totalService.get(email) + serviceAmt);
    }

    // Grab the names from the roster
    // key is email, values is name
    let emailMap: Map<string, string> = new Map();
    let rosterVals = rosterTable.getRange().getValues();
    for (let n = 7; n < rosterVals.length; n++) {
        emailMap.set(rosterVals[n][4].toString(), rosterVals[n][1].toString());
        // have to hard code 4 and 1
    }

    let currentRow: number = 2
    totalService.forEach((val, key) => {
        let tempRange = totalSheet.getRange("A" + currentRow);
        tempRange.setValue(key)
        
        tempRange = totalSheet.getRange("B" + currentRow);
        tempRange.setValue(emailMap.get(key));

        tempRange = totalSheet.getRange("C" + currentRow);
        tempRange.setValue(val);
        currentRow += 1;
    });
}
