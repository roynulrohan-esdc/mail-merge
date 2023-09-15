import { get, readable, writable } from "svelte/store";
import { path } from "./settings";
import { pageLoading } from './routes'
import { config } from "./emails";

export const dataFilesList = writable([]);
export const dataError = writable()
export const data = writable();

const subDirectory = '\\input\\data';

export const dataDirectory = path + subDirectory;

const errors = {
  1: { message: "Input directory does not exist.", path: dataDirectory },
  2: { message: "No data files found", path: dataDirectory }
}

export const loadData = async () => {
  try {
    let fso = new ActiveXObject("Scripting.FileSystemObject");

    if (!fso.FolderExists(dataDirectory)) {
      throw 1
    }

    let dataFolder = fso.GetFolder(dataDirectory);

    const dataFiles = readFileNames(new Enumerator(dataFolder.files));

    if (dataFiles.length === 0) {
      throw 2
    }

    dataFilesList.set(dataFiles)
  } catch (e) {
    console.log(e)
    dataError.set(errors[e] || e)
  }
};

const readFileNames = (enumerator) => {
  const list = [];

  for (; !enumerator.atEnd(); enumerator.moveNext()) {
    console.log(enumerator.item().Type)
    if (enumerator.item().Type === 'Microsoft Excel Worksheet') {
      list.push(enumerator.item().Name);
    }
  }

  return list;
}

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

export const importData = async (filePath) => {
  filePath = dataDirectory + "\\" + filePath;
  pageLoading.set(true);

  setTimeout(async () => {
    try {
      data.set(await readDataFile(filePath))
    } catch (e) {
      console.log(e)
    }

    pageLoading.set(false)
  }, 1000);
}

export const readDataFile = async (filePath) => {
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
      const classification = currentArea.Cells(j, 6).value;

      if (lastName && firstName && email && classification) {

        array.push({ lastName, firstName, email, classification });
      }
    }
  }

  return array;
};
