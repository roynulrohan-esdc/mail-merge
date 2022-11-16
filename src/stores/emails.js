import { get, writable } from "svelte/store"
import { path } from "./settings";
import { data } from './data'
import { templatesDirectory } from './templates'

export const employeeEmails = writable([])
export const managerEmails = writable([])

export const sendingEmails = writable(false);
export const sendingMessage = writable("");

export const generatingEmails = writable(false);
export const generationMessage = writable("");

const outputPath = path + '/output/';

export const sendEmails = (mode = 0) => {
    sendingEmails.set(true);
    sendingMessage.set("");

    setTimeout(() => {
        if (mode === 0) {
            try {
                get(employeeEmails).forEach((email) => {
                    email.Send()
                })
            } catch (e) {
                console.error(e)
                sendingMessage.set({ message: "An unknown error occured while sending emails" })
            }
        } else {
            try {
                get(managerEmails).forEach((email) => {
                    email.Send()
                })
            } catch (e) {
                console.error(e)
                sendingMessage.set({ message: "An unknown error occured while sending emails" })
            }
        }

        sendingMessage.set({ message: "Emails successfully sent" })
        sendingEmails.set(false);
    }, 1000);
}

export const generateEmails = (mode = 0, templateName) => {
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

            const employees = [...get(data)];

            const emails = []

            employees.forEach((employee) => {
                try {
                    let objOutlook = new ActiveXObject("Outlook.Application")

                    let objEmail = objOutlook.CreateItem(0) //0 is email

                    objEmail.To = employee.email;

                    const email = getEmployeeEmail(templateName, employee);

                    objEmail.Subject = email.subject

                    objEmail.HTMLBody = email.body;

                    objEmail.SaveAs(`${outputPath}/employees/${employee.fullName.split(" ")[1]}, ${employee.fullName.split(" ")[0]} - ${email.subject}.msg`)

                    emails.push(objEmail)
                }

                catch (e) {
                    console.error(e)
                    generationMessage.set({ message: "An unknown error occured while generating emails" })
                }
            })

            employeeEmails.set(emails)

            generationMessage.set({ message: 'Employee emails succesfully generated at ', path: '/output/employees/' })
        } else {
            if (!fso.FolderExists(outputPath + "/managers")) {
                fso.CreateFolder(outputPath + "/managers");
            } else {
                fso.DeleteFolder(outputPath + "/managers");
                fso.CreateFolder(outputPath + "/managers");
            }

            const employees = [...get(data)];
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

            const emails = []

            managers.forEach((manager) => {
                try {
                    let objOutlook = new ActiveXObject("Outlook.Application")

                    let objEmail = objOutlook.CreateItem(0) //0 is email

                    objEmail.To = managerToEmail[manager[0]];

                    const email = getManagerEmail(templateName, { supervisorName: manager[0], employees: manager[1] });

                    objEmail.Subject = email.subject

                    objEmail.HTMLBody = email.body

                    objEmail.SaveAs(`${outputPath}/managers/${manager[0].split(" ")[1]}, ${manager[0].split(" ")[0]} - ${email.subject}.msg`)

                    emails.push(objEmail)
                }

                catch (e) {
                    console.error(e)
                    generationMessage.set({ message: "An unknown error occured while generating emails" })
                }
            })

            managerEmails.set(emails)

            generationMessage.set({ message: 'Manager emails succesfully generated at ', path: '/output/managers/' })
        }

        generatingEmails.set(false)
    }, 1000);
}


const getEmployeeEmail = (templateName, { fullName }) => {
    let outlook = new ActiveXObject("Outlook.Application");

    let mailItem = outlook.createItemFromTemplate(templatesDirectory + "\\employee\\" + templateName)

    const keys = {
        fullName,
    }

    let body = mailItem.HTMLBody;

    Object.keys(keys).forEach(key => {
        body = body.replaceAll(`{${key}}`, keys[key])
    })

    return { body, subject: mailItem.Subject }
}

const getManagerEmail = (templateName, { supervisorName, employees }) => {
    const formattedEmployees = employees.map((employee) => {
        return `<li style="${styles.email}">${employee}</li>`
    })

    let outlook = new ActiveXObject("Outlook.Application");

    let mailItem = outlook.createItemFromTemplate(templatesDirectory + "\\manager\\" + templateName)

    const keys = {
        supervisorName, employees: formattedEmployees.join("")
    }

    let body = mailItem.HTMLBody;

    Object.keys(keys).forEach(key => {
        body = body.replaceAll(`{${key}}`, keys[key])
    })

    return { body, subject: mailItem.Subject }
}

const styles = {
    email: `font-size:12pt; font-family:Arial;`
}