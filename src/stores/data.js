import { get, readable, writable } from "svelte/store";
import { path } from "./settings";
import { pageLoading } from './routes'
import { config } from "./emails";


export const data = writable();
export const failLoadData = writable(false);
export const failLoadMessage = writable('');

export const FILE_PATH = "\\input\\Data.xlsx"

export const getFileName = () => {
  const split = FILE_PATH.split('\\');

  return split[split.length - 1]
}

export const loadData = async () => {
  try {
    data.set(readFile(path + FILE_PATH))
  } catch (e) {
    console.log(e)
    failLoadMessage.set(e)
    failLoadData.set(true)
  }
};

export const openFile = () => {
  try {
    readFileWithVisibility(path + FILE_PATH)
    pageLoading.set(false)
  } catch (e) {
    console.log(e)
  }
}

let workbook;

const readFileWithVisibility = (filePath) => {
  let fso = new ActiveXObject("Scripting.FileSystemObject");

  if (fso.FileExists(filePath)) {
    const DEFAULT_DPI = 96;

    let excel = new ActiveXObject("Excel.Application");
    excel.Visible = true
    excel.Left = (screen.availWidth / (screen.systemXDPI / DEFAULT_DPI * 2)) * 0.75
    excel.Top = 0
    excel.Width = (screen.availWidth / (screen.systemXDPI / DEFAULT_DPI * 2)) * 0.75
    excel.Height = (screen.availHeight / (screen.systemXDPI / DEFAULT_DPI)) * 0.75
    excel.DisplayAlerts = false

    window.addEventListener(
      "beforeunload",
      function (e) {
        excel.DisplayAlerts = false;
        excel.quit();
      },
      false
    );

    try {
      workbook = excel.workbooks.open(filePath);
      workbook.activate();
    } catch (e) {
      console.error(e);
      throw `Error opening ${filePath}`
    }
  } else {
    throw `${filePath} does not exist`
  }
}

const readFile = (filePath) => {
  let fso = new ActiveXObject("Scripting.FileSystemObject");
  let array;

  if (fso.FileExists(filePath)) {
    let excel = new ActiveXObject("Excel.Application");
    excel.Visible = false;

    window.addEventListener(
      "beforeunload",
      function (e) {
        excel.DisplayAlerts = false;
        excel.quit();
      },
      false
    );

    try {
      let workbook = excel.workbooks.open(filePath);
      workbook.activate();

      array = readData(workbook);

    } catch (e) {
      console.error(e);
      throw `Error opening excel`
    }
    excel.quit();
  } else {
    throw `${filePath} does not exist`
  }

  return array;
};

export const readData = (workbook) => {
  let sheet;

  try {
    sheet = workbook.Worksheets(get(config).sheet);
  } catch (e) {
    throw "Error reading excel. Invalid sheet name"
  }

  let remainingResults = sheet.UsedRange.SpecialCells(12);
  let areasResults = remainingResults.Areas;

  const array = [];

  for (let i = 1; i <= remainingResults.Areas.Count; i++) {
    let currentArea = areasResults(i);
    let areaRows = currentArea.Rows.Count;

    for (let j = 2; j <= areaRows; j++) {
      const lastName = currentArea.Cells(j, 1).value;
      const firstName = currentArea.Cells(j, 2).value;
      const email = currentArea.Cells(j, 3).value;
      const costCentre = currentArea.Cells(j, 4).value;
      const manager = currentArea.Cells(j, 5).value;
      const managerClassification = currentArea.Cells(j, 6).value;
      const managerEmail = currentArea.Cells(j, 7).value;

      if (lastName && firstName && email && costCentre && manager && managerClassification && managerEmail) {

        array.push({ lastName, firstName, email, costCentre, manager, managerClassification, managerEmail });
      }
    }
  }

  return array;
};
