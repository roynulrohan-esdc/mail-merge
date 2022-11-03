import { get, writable } from "svelte/store"
import { path } from "./settings";
import { data } from './data'

export const generatingEmails = writable(false);
export const generationMessage = writable("");

const outputPath = path + '/output/';

export const generateEmails = (mode = 0) => {
    let fso = new ActiveXObject("Scripting.FileSystemObject");

    if (!fso.FolderExists(outputPath)) {
        fso.CreateFolder(outputPath);
    }

    generatingEmails.set(true);
    generationMessage.set("");

    setTimeout(() => {
        if (mode === 0) {
            if (!fso.FolderExists(outputPath + "/employees")) {
                fso.CreateFolder(outputPath + "/employees");
            } else {
                fso.DeleteFolder(outputPath + "/employees");
                fso.CreateFolder(outputPath + "/employees");
            }

            const employees = [...get(data).scenarioOne];

            employees.forEach((employee) => {
                try {
                    let objOutlook = new ActiveXObject("Outlook.Application")

                    let objEmail = objOutlook.CreateItem(0) //0 is email

                    objEmail.To = employee.email;

                    objEmail.Subject = "Test Subject"

                    objEmail.HTMLBody = scenarioOneEmailBody(employee.firstName, employee.lastName);

                    objEmail.SaveAs(`${outputPath}/employees/${employee.lastName}, ${employee.firstName} - Test Subject.msg`)
                }

                catch (e) {
                    console.error(e)
                    generationMessage.set("An unknown error occured while generating emails")
                }
            })

            generatingEmails.set(false)
            generationMessage.set({ message: 'Employee emails succesfully generated at ', path: '/output/employees/' })
        } else {
            if (!fso.FolderExists(outputPath + "/managers")) {
                fso.CreateFolder(outputPath + "/managers");
            } else {
                fso.DeleteFolder(outputPath + "/managers");
                fso.CreateFolder(outputPath + "/managers");
            }

            const employees = [...get(data).scenarioTwo];
            const managerToEmployees = {};
            const managerToEmail = {};

            employees.forEach((employee) => {
                if (!managerToEmployees[employee.supervisorName]) {
                    managerToEmployees[employee.supervisorName] = [];
                }

                managerToEmployees[employee.supervisorName].push(employee.fullName);

                if (!managerToEmail[employee.supervisorName]) {
                    managerToEmail[employee.supervisorName] = employee.supervisorEmail;
                }
            })

            const managers = Object.entries(managerToEmployees);

            managers.forEach((manager) => {
                try {
                    let objOutlook = new ActiveXObject("Outlook.Application")

                    let objEmail = objOutlook.CreateItem(0) //0 is email

                    objEmail.To = managerToEmail[manager[0]];

                    objEmail.Subject = "Test Subject"

                    objEmail.HTMLBody = scenarioTwoEmailBody(manager[0], manager[1]);

                    objEmail.SaveAs(`${outputPath}/managers/${manager[0].split(" ")[1]}, ${manager[0].split(" ")[0]} - Test Subject.msg`)
                }

                catch (e) {
                    console.error(e)
                    generationMessage.set("An unknown error occured while generating emails")
                }
            })

            generatingEmails.set(false)
            generationMessage.set({ message: 'Manager emails succesfully generated at ', path: '/output/managers/' })
        }
    }, 200);
}


const scenarioOneEmailBody = (firstName, lastName) => {
    let body = `

    <p style="${styles.email}">
        Hello ${firstName} ${lastName},                                                         <br/>

                                                                                                <br/>

        This is a test email.                                                                   <br/>
    </p>`


    return body
}

const scenarioTwoEmailBody = (fullName, employees) => {
    const formattedEmployees = employees.map((employee) => {
        return `<li style="${styles.email}">${employee}</li>`
    })
    let body = `

    <p style="${styles.email}">
        Hello ${fullName},                                                                      <br/>

                                                                                                <br/>

        This is a test email regarding the following employees:                                 <br/>
        <ul>${formattedEmployees.join("\n")}</ul>                                               <br/>
    </p>`


    return body
}

const styles = {
    email: `font-size:12pt; font-family:Arial;`
}