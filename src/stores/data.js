import { get, readable, writable } from "svelte/store";
import { path } from "./settings";

export const data = writable({ scenarioOne: [], scenarioTwo: [] });
export const failLoadData = writable(false);

const DEFAULT_SHEET = "Sheet1";

export const loadData = () => {
  readFile(path + "/input/Scenario1.xlsx", 1);
  readFile(path + "/input/Scenario2.xlsx", 2);
};

const readFile = (filePath, scenario) => {
  let fso = new ActiveXObject("Scripting.FileSystemObject");

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
        readScenarioOne(workbook);
      } else {
        readScenarioTwo(workbook);
      }

      failLoadData.set(false);
    } catch (e) {
      console.error(e);
      failLoadData.set(true);
    }

    excel.quit();
  } else {
    failLoadData.set(true);
  }
};

export const readScenarioOne = (workbook) => {
  const sheet = workbook.Worksheets(DEFAULT_SHEET);

  let remainingResults = sheet.UsedRange.SpecialCells(12);
  let areasResults = remainingResults.Areas;

  for (let i = 1; i <= remainingResults.Areas.Count; i++) {
    let currentArea = areasResults(i);
    let areaRows = currentArea.Rows.Count;

    for (let j = 2; j <= areaRows; j++) {
      const firstName = currentArea.Cells(j, 1).value;
      const lastName = currentArea.Cells(j, 2).value;
      const email = currentArea.Cells(j, 3).value;

      get(data).scenarioOne.push({ firstName, lastName, email });
    }
  }
};

export const readScenarioTwo = (workbook) => {
  const sheet = workbook.Worksheets(DEFAULT_SHEET);

  let remainingResults = sheet.UsedRange.SpecialCells(12);
  let areasResults = remainingResults.Areas;

  for (let i = 1; i <= remainingResults.Areas.Count; i++) {
    let currentArea = areasResults(i);
    let areaRows = currentArea.Rows.Count;

    for (let j = 2; j <= areaRows; j++) {
      const fullName = currentArea.Cells(j, 1).value;
      const email = currentArea.Cells(j, 2).value;
      const supervisorName = currentArea.Cells(j, 3).value;
      const supervisorEmail = currentArea.Cells(j, 4).value;

      get(data).scenarioTwo.push({ fullName, email, supervisorName, supervisorEmail });
    }
  }
};
