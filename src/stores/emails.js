import { get, writable } from "svelte/store"
import { path } from "./settings";
import { data } from './data'
import { templatesDirectory } from './templates'
import configJson from '../../public/input/mailConfig.json';

export const employeeEmails = writable([])
export const managerEmails = writable([])

export const sendingEmails = writable(false);
export const sendingMessage = writable("");

export const generatingEmails = writable(false);
export const generationMessage = writable("");

export const configError = writable();
export const config = writable({ mailbox: null })

const outputPath = path + '/output/';

export const loadMailConfig = () => {
    if (!configJson.mailbox) {
        configError.set('Attribute "mailbox" not found in \\input\\mailConfig.json')
    }

    config.set({ mailbox: configJson.mailbox })
}

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

        sendingMessage.set({ message: "Emails successfully forwarded to outlook" })
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

                    objEmail.SentOnBehalfOfName = get(config).mailbox

                    const email = getEmployeeEmail(templateName, employee);

                    objEmail.Subject = email.subject

                    objEmail.HTMLBody = email.body;

                    objEmail.SaveAs(`${outputPath}/employees/${employee.lastName}, ${employee.firstName}.msg`)

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
                const managers = employee.managers.split(',');
                const managersEmails = employee.managersEmails.split(',')

                managers.forEach((manager, i) => {
                    manager = manager.trim();

                    if (!managerToEmployees[manager]) {
                        managerToEmployees[manager] = [];
                    }

                    managerToEmployees[manager].push(`${employee.firstName} ${employee.lastName}`);

                    if (!managerToEmail[manager]) {
                        managerToEmail[manager] = managersEmails[i].trim();
                    }
                })


            })

            const managers = Object.entries(managerToEmployees);

            const emails = []

            managers.forEach((manager) => {
                try {
                    let objOutlook = new ActiveXObject("Outlook.Application")

                    let objEmail = objOutlook.CreateItem(0) //0 is email

                    objEmail.To = managerToEmail[manager[0]];
                    
                    objEmail.SentOnBehalfOfName = get(config).mailbox

                    const email = getManagerEmail(templateName, { managerName: manager[0], impactedEmployees: manager[1] });

                    objEmail.Subject = email.subject

                    objEmail.HTMLBody = email.body

                    objEmail.SaveAs(`${outputPath}/managers/${manager[0].split(" ")[1]}, ${manager[0].split(" ")[0]}.msg`)

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


const getEmployeeEmail = (templateName, { firstName, lastName }) => {
    let outlook = new ActiveXObject("Outlook.Application");

    let mailItem = outlook.createItemFromTemplate(templatesDirectory + "\\employee\\" + templateName)

    const keys = {
        firstName, lastName
    }

    let body = mailItem.HTMLBody;

    Object.keys(keys).forEach(key => {
        body = body.replaceAll(`{${key}}`, keys[key])
    })

    return { body, subject: mailItem.Subject }
}

const getManagerEmail = (templateName, { managerName, impactedEmployees }) => {
    const formattedEmployees = impactedEmployees.map((employee) => {
        return `${employee}`
    })

    let outlook = new ActiveXObject("Outlook.Application");

    let mailItem = outlook.createItemFromTemplate(templatesDirectory + "\\manager\\" + templateName)

    const keys = {
        managerName, impactedEmployees: `${formattedEmployees.join(', ')}`
    }

    let body = mailItem.HTMLBody;

    Object.keys(keys).forEach(key => {
        body = body.replaceAll(`{${key}}`, keys[key])
    })

    console.log(body)

    return { body, subject: mailItem.Subject }
}

const styles = {
    email: `font-size:12pt; font-family:Arial;`
}