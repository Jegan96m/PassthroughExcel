function convertToXlsx() {
    const csvFile = document.getElementById('csvFile').files[0];
    if (!csvFile) {
        alert('Please select a CSV file.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const csvData = e.target.result;
        const data = CSVToArray(csvData);

        const headers = data[0];
        const rows = data.slice(1);

        const sheets = {};
        const workbook = XLSX.utils.book_new();
        let workbookName = '';
        let bookName='';

        let sheetNumber=1;

        rows.forEach(row => {
            const rowData = {};
            headers.forEach((header, index) => {
                rowData[header] = row[index];
            });

            if (!workbookName) {
                workbookName = rowData['Zone'];
            }

            if (!bookName) {
                bookName = rowData['v_project'];
            }

            const vPlanValue = rowData['v_plan'];
            if (vPlanValue !== undefined) { // Only add rows with defined v_plan
                if (!sheets[vPlanValue]) {
                    sheets[vPlanValue] = [];
                }

                const selectedRow = {
                    'ID': rowData['ID'],
                    'Size': rowData['Size'],
                    'Zone': rowData['Zone'],
                    'Service Group': rowData['Service Group'],
                    'Service Area': rowData['Service Area'],
                    'Tier Rating': rowData['Tier Rating'],
                    '# of Primary Splitters': rowData['# of Primary Splitters'],
                    '# of Secondary Splitters': rowData['# of Secondary Splitters'],
                    'v_created_time': rowData['v_created_time'],
                    'v_last_edited_time': rowData['v_last_edited_time'],
                    'Passthrough': [],
                    'Producer':  [],
                    'Complete Date':  [],
                    
                };

                sheets[vPlanValue].push(selectedRow);
            }
        });

        
    
        // Create summary sheet with sheet names and additional headers
        const summaryData = Object.keys(sheets).map(sheetName => ({
            
            'No':sheetNumber++,
            'SG': sheetName,
            'Number of HH': `=COUNTA('${sheetName}'!$A$2:$A1000)`,
            'Number of Y HH': `=COUNTIF('${sheetName}'!$K$2:$K1000,"Y")`, // Formula to count
            'Number of N HH':  `=COUNTIF('${sheetName}'!$K$2:$K1000,"N")`,
            'Remaining':`=IF(ISBLANK($C${sheetNumber}),"No Handhole",IF($C${sheetNumber}-($D${sheetNumber}+$E${sheetNumber})=0,"COMPLETED",$C${sheetNumber}-($D${sheetNumber}+$E${sheetNumber})))`,
            'Producer': `=UPPER(IFNA(INDEX('${sheetName}'!$L$2:$L1000,MATCH(TRUE,LEN('${sheetName}'!$L$2:$L1000)>1,0)),""))`,
            'Remark':'',
            
        }));
        
        

        // Add a "Total" row after the last plan in summary sheet
        summaryData.push({
            'SG': 'Total',
            'Number of HH': `=SUM(C2:C${summaryData.length + 1})`,
            'Number of Y HH': `=SUM(D2:D${summaryData.length + 1})`,
            'Number of N HH': `=SUM(E2:E${summaryData.length + 1})`,
            'Remaining': `=SUM(F2:F${summaryData.length + 1})`
            
        });

        const summarySheet = XLSX.utils.json_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(workbook, summarySheet, 'Overview');

        Object.keys(sheets).forEach(sheetName => {
            const sheetData = sheets[sheetName];
            const ws = XLSX.utils.json_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(workbook, ws, sheetName);
        });

        const xlsxFileName = `${bookName} handhole status.xlsx`;
        XLSX.writeFile(workbook, xlsxFileName);

        //alert(`Conversion successful! The file has been saved as "${xlsxFileName}".`);
    };
    reader.readAsText(csvFile);
}

function CSVToArray(strData, strDelimiter = ",") {
    const objPattern = new RegExp(
        (
            "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
            "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
            "([^\"\\" + strDelimiter + "\\r\\n]*))"
        ),
        "gi"
    );

    const arrData = [[]];
    let arrMatches = null;

    while (arrMatches = objPattern.exec(strData)) {
        const strMatchedDelimiter = arrMatches[1];
        if (strMatchedDelimiter.length && strMatchedDelimiter !== strDelimiter) {
            arrData.push([]);
        }

        const strMatchedValue = arrMatches[2] ?
            arrMatches[2].replace(new RegExp("\"\"", "g"), "\"") :
            arrMatches[3];

        arrData[arrData.length - 1].push(strMatchedValue);
    }

    return arrData;
}
