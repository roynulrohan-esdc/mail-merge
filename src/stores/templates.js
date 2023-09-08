import { writable } from "svelte/store";
import { path } from "./settings";

export const templatesList = writable({ employees: [], managers: [] });

export const templatesError = writable()

const subDirectory = '\\input\\scripts';

export const templatesDirectory = path + subDirectory;

const errors = {
    1: { message: "Scripts directory does not exist.", path: templatesDirectory },
    2: { message: "Employee Scripts directory does not exist.", path: subDirectory + '\\employee' },
    3: { message: "Manager Scripts directory does not exist.", path: subDirectory + '\\manager' },
    4: { message: "Employee Scripts directory empty. At least one Outlook Template must be present.", path: subDirectory + '\\employee' },
    5: { message: "Manager Scripts directory empty. At least one Outlook Template must be present.", path: subDirectory + '\\manager' }
}

export const loadTemplates = async () => {
    try {
        let fso = new ActiveXObject("Scripting.FileSystemObject");

        if (!fso.FolderExists(templatesDirectory)) {
            throw 1
        }

        if (!fso.FolderExists(templatesDirectory + '\\employee')) {
            throw 2
        }

        let employeeFolder = fso.GetFolder(templatesDirectory + '\\employee');

        const employees = readFileNames(new Enumerator(employeeFolder.files));

        templatesList.set({ employees, manager:[] })
    } catch (e) {
        console.log(e)
        templatesError.set(errors[e] || e)
    }
};

const readFileNames = (enumerator) => {
    const list = [];

    for (; !enumerator.atEnd(); enumerator.moveNext()) {
        if (enumerator.item().Type === 'Outlook Item') {
            list.push(enumerator.item().Name);
        }
    }

    return list;
}
