const fs = require('fs');
const XLSX = require('xlsx');

function compareNameAndStatus(file1, file2, nameCol = 'Name', statusCol = 'Status', outputFileName = 'name_status_comparison.xlsx') {
    // Read Excel files
    const workbook1 = XLSX.readFile(file1);
    const workbook2 = XLSX.readFile(file2);

    // Extract sheet names assuming first sheet in both files
    const sheetName1 = workbook1.SheetNames[0];
    const sheetName2 = workbook2.SheetNames[0];

    // Convert Excel sheets to JSON objects
    const sheet1 = XLSX.utils.sheet_to_json(workbook1.Sheets[sheetName1]);
    const sheet2 = XLSX.utils.sheet_to_json(workbook2.Sheets[sheetName2]);

    // Extract columns for comparison
    const data1 = sheet1.map(row => ({ [nameCol]: row[nameCol], [statusCol]: row[statusCol] }));
    const data2 = sheet2.map(row => ({ [nameCol]: row[nameCol], [statusCol]: row[statusCol] }));

    // Create a set of names from both files
    const names1 = new Set(data1.map(row => row[nameCol]));
    const names2 = new Set(data2.map(row => row[nameCol]));

    // Find the intersection of the sets
    const commonNames = new Set([...names1].filter(name => names2.has(name)));

    // Filter data arrays to only include rows with common names
    const filteredData1 = data1.filter(row => commonNames.has(row[nameCol]));
    const filteredData2 = data2.filter(row => commonNames.has(row[nameCol]));

    // Perform inner join on Name column with filtered data
    const mergedData = filteredData1.reduce((acc, obj1) => {
        const obj2 = filteredData2.find(obj => obj[nameCol] === obj1[nameCol]);
        if (obj2) {
            acc.push({
                [`${nameCol}_file1`]: obj1[nameCol],
                [`${nameCol}_file2`]: obj2[nameCol],
                [`${statusCol}_file1`]: obj1[statusCol],
                [`${statusCol}_file2`]: obj2[statusCol]
            });
        }
        return acc;
    }, []);

    // Convert result data to Excel workbook
    const ws = XLSX.utils.json_to_sheet(mergedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Comparison Result');

    // Write to output file
    XLSX.writeFile(wb, outputFileName);

    console.log(`Comparison of '${statusCol}' for the same '${nameCol}' has been written to '${outputFileName}'.`);
}

// Specify input and output files
const file1 = '/Users/sonali.kashyap/Downloads/AWS harness (1).xlsx';
const file2 = '/Users/sonali.kashyap/Downloads/GCP SMOKE AV2.xlsx';
const outputFileName = 'DIFF REPORT.xlsx';

// Specify name and status columns
const nameCol = 'Name';
const statusCol = 'Status';

// Run comparison function
compareNameAndStatus(file1, file2, nameCol, statusCol, outputFileName);
