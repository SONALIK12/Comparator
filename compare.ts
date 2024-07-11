import * as XLSX from 'xlsx';
import _ from 'lodash';

interface DataFrame {
  [key: string]: any[];
}

const readExcel = (filePath: string): DataFrame => {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1 });
};

const writeExcel = (data: any[], filePath: string): void => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, filePath);
};

const compareNameAndStatus = (
  file1: string,
  file2: string,
  nameCol: string = 'Name',
  statusCol: string = 'Status',
  outputFile: string = 'name_status_comparison.xlsx'
): void => {
  const df1 = readExcel(file1);
  const df2 = readExcel(file2);

  const header1 = df1[0];
  const header2 = df2[0];

  if (!header1.includes(nameCol) || !header2.includes(nameCol)) {
    console.error(`Column '${nameCol}' must be present in both files.`);
    return;
  }
  if (!header1.includes(statusCol) || !header2.includes(statusCol)) {
    console.error(`Column '${statusCol}' must be present in both files.`);
    return;
  }

  const df1Rows = df1.slice(1);
  const df2Rows = df2.slice(1);

  const nameIndex1 = header1.indexOf(nameCol);
  const statusIndex1 = header1.indexOf(statusCol);
  const nameIndex2 = header2.indexOf(nameCol);
  const statusIndex2 = header2.indexOf(statusCol);

  const df1Map = _.keyBy(df1Rows, row => row[nameIndex1]);
  const df2Map = _.keyBy(df2Rows, row => row[nameIndex2]);

  const mergedData = _.intersection(Object.keys(df1Map), Object.keys(df2Map)).map(name => ({
    [`${nameCol}_file1`]: name,
    [`${statusCol}_file1`]: df1Map[name][statusIndex1],
    [`${nameCol}_file2`]: name,
    [`${statusCol}_file2`]: df2Map[name][statusIndex2]
  }));

  writeExcel(mergedData, outputFile);

  console.log(`Comparison of '${statusCol}' for the same '${nameCol}' has been written to '${outputFile}'.`);
};

const file1 = '/Users/sonali.kashyap/Downloads/SMOKE AWS (1).xlsx';
const file2 = '/Users/sonali.kashyap/Downloads/SMOKE GCP (1).xlsx';
const nameCol = 'Name';
const statusCol = 'Status';
const outputFile = 'name_status_comparison.xlsx';

compareNameAndStatus(file1, file2, nameCol, statusCol, outputFile);
