import { get, readable, writable } from "svelte/store";
import { path } from "./settings";

export const data = writable();
export const failLoadData = writable(false);
export const failLoadMessage = writable('');

const DEFAULT_SHEET = "Sheet1";

export const loadData = () => {
  try {
    const scenarioOne = readFile(path + "/input/Scenario1.xlsx", 1);
    const scenarioTwo = readFile(path + "/input/Scenario2.xlsx", 2);

    data.set({ scenarioOne, scenarioTwo })
  } catch (e) {
    console.log(e)
    failLoadMessage.set(e)
    failLoadData.set(true)
  }
};

const readFile = (filePath, scenario) => {
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


      if (scenario == 1) {
        array = readScenarioOne(workbook);
      } else {
        array = readScenarioTwo(workbook);
      }
    } catch (e) {
      console.error(e);
      throw `Error opening ${filePath}`
    }
    excel.quit();
  } else {
    throw `${filePath} does not exist`
  }

  return array;
};

export const readScenarioOne = (workbook) => {
  const sheet = workbook.Worksheets(DEFAULT_SHEET);

  let remainingResults = sheet.UsedRange.SpecialCells(12);
  let areasResults = remainingResults.Areas;

  const array = [];

  for (let i = 1; i <= remainingResults.Areas.Count; i++) {
    let currentArea = areasResults(i);
    let areaRows = currentArea.Rows.Count;

    for (let j = 2; j <= areaRows; j++) {
      const firstName = currentArea.Cells(j, 1).value;
      const lastName = currentArea.Cells(j, 2).value;
      const email = currentArea.Cells(j, 3).value;

      if (firstName && lastName && email) {
        array.push({ firstName, lastName, email });
      }
    }
  }

  return array;
};

export const readScenarioTwo = (workbook) => {
  const sheet = workbook.Worksheets(DEFAULT_SHEET);

  let remainingResults = sheet.UsedRange.SpecialCells(12);
  let areasResults = remainingResults.Areas;

  const array = [];

  for (let i = 1; i <= remainingResults.Areas.Count; i++) {
    let currentArea = areasResults(i);
    let areaRows = currentArea.Rows.Count;

    for (let j = 2; j <= areaRows; j++) {
      const fullName = currentArea.Cells(j, 1).value;
      const email = currentArea.Cells(j, 2).value;
      const supervisorName = currentArea.Cells(j, 3).value;
      const supervisorEmail = currentArea.Cells(j, 4).value;

      if (fullName && email && supervisorName && supervisorEmail) {
        array.push({ fullName, email, supervisorName, supervisorEmail });
      }
    }
  }

  return array;
};
