import { get, writable } from "svelte/store";
import { path } from "./settings";
import { data } from "./data";
import { templatesDirectory } from "./templates";
import FileSystem from "../utility/filesystem";
import { TemplateError, createTemplate } from "../utility/templater";

export const employeeEmails = writable([]);
export const managerEmails = writable([]);

export const sendingEmails = writable(false);
export const sendingMessage = writable({ message: "", path: "" });

export const emailsSent = writable(false);

export const generatingEmails = writable(false);
export const generationMessage = writable({ message: "", path: "" });
export const generationCount = writable(0);

export const configError = writable();
export const config = writable({ mailbox: null, sheet: null });

const outputPath = path + "/output/";

export const loadMailConfig = () => {
  const fso = new ActiveXObject("Scripting.FileSystemObject");

  try {
    const file = fso.OpenTextFile(path + "/input/mailConfig.json");
    const json = JSON.parse(file.ReadAll());

    config.set({ ...json });

    file.close();
    if (!json.sheet && !json.mailbox) {
      configError.set("Error loading mail configuration");
    }
  } catch (e) {
    configError.set(e);
    console.error(e);
  }
};

export const sendEmails = (mode = 0) => {
  sendingEmails.set(true);
  sendingMessage.set({ message: "", path: "" });

  setTimeout(() => {
    if (mode === 0) {
      try {
        const emails = [...get(employeeEmails)];

        for (let index = 0; index < emails.length; index++) {
          const email = emails[index];

          email.send();
        }
      } catch (e) {
        console.error(e);
        sendingMessage.set({ message: "An unknown error occured while sending emails", path: "" });
      }
    } else {
      try {
        get(managerEmails).forEach((email) => {
          email.Send();
        });
      } catch (e) {
        console.error(e);
        sendingMessage.set({ message: "An unknown error occured while sending emails", path: "" });
      }
    }

    emailsSent.set(true);
    sendingMessage.set({ message: "Emails successfully forwarded to outlook", path: "" });
    sendingEmails.set(false);
  }, 1000);
};

export const generateEmails = (mode = 0, templateName) => {
  let fso = new ActiveXObject("Scripting.FileSystemObject");

  if (!fso.FolderExists(outputPath)) {
    fso.CreateFolder(outputPath);
  }

  generatingEmails.set(true);
  generationMessage.set({ message: "", path: "" });

  setTimeout(async () => {
    if (mode === 0) {
      if (!fso.FolderExists(outputPath + "/employees")) {
        fso.CreateFolder(outputPath + "/employees");
      } else {
        fso.DeleteFolder(outputPath + "/employees");
        fso.CreateFolder(outputPath + "/employees");
      }

      const employees = [...(get(data) as any).employees];

      const emails = [];

      for (let index = 0; index < employees.length; index++) {
        const employee = employees[index];

        const employeeDirectory = outputPath + "/employees/" + `${employee.lastName}, ${employee.firstName}/`;

        FileSystem.createFolder(employeeDirectory);

        const template = createTemplate(path + "\\input\\attachments\\CDS Welcome Letter Template BIL.docx", employee);

        if (template === TemplateError.INVALID_TEMPLATE) {
          generationMessage.set({ message: "An unknown error occured while generating emails", path: "" });
          return;
        }
        if (template === TemplateError.MISSING_LOCKED_FILE) {
          generationMessage.set({ message: "An unknown error occured while generating emails", path: "" });
          return;
        }
        if (template === TemplateError.UNKNOWN_ERROR) {
          generationMessage.set({ message: "An unknown error occured while generating emails", path: "" });
          return;
        }

        await FileSystem.saveFile(template, employeeDirectory + "CDS Welcome Letter Template BIL.docx", true);

        try {
          let objOutlook = new ActiveXObject("Outlook.Application");

          let objEmail = objOutlook.CreateItem(0); //0 is email

          const email = getEmployeeEmail(templateName, employee);

          objEmail = email;

          objEmail.To = employee.email;

          objEmail.SentOnBehalfOfName = get(config).mailbox;

          objEmail.Attachments.Add(employeeDirectory + "CDS Welcome Letter Template BIL.docx");

          objEmail.SaveAs(employeeDirectory + `${employee.lastName}, ${employee.firstName}.msg`);

          emails.push(objEmail);
        } catch (e) {
          console.error(e);
          generationMessage.set({ message: "An unknown error occured while generating emails", path: "" });
        }
      }

      employeeEmails.set(emails);

      generationMessage.set({ message: "Employee emails succesfully generated", path: "" });
    }
    // else {
    //     if (!fso.FolderExists(outputPath + "/managers")) {
    //         fso.CreateFolder(outputPath + "/managers");
    //     } else {
    //         fso.DeleteFolder(outputPath + "/managers");
    //         fso.CreateFolder(outputPath + "/managers");
    //     }

    //     const employees = [...get(data).employees];
    //     const managerToEmployees = {};
    //     const managerToEmail = {};

    //     employees.forEach((employee) => {
    //         const manager = employee.manager.trim();
    //         const managerEmail = employee.managerEmail.trim();

    //         if (!managerToEmployees[manager]) {
    //             managerToEmployees[manager] = [];
    //         }

    //         managerToEmployees[manager].push(`${employee.firstName} ${employee.lastName}`);

    //         if (!managerToEmail[manager]) {
    //             managerToEmail[manager] = managerEmail.trim();
    //         }

    //     })

    //     const managers = Object.entries(managerToEmployees);

    //     const emails = []

    //     managers.forEach((manager) => {
    //         try {
    //             let objOutlook = new ActiveXObject("Outlook.Application")

    //             let objEmail = objOutlook.CreateItem(0) //0 is email

    //             objEmail.To = managerToEmail[manager[0]];

    //             objEmail.SentOnBehalfOfName = get(config).mailbox

    //             const email = getManagerEmail(templateName, { impactedEmployees: manager[1] });

    //             objEmail.Subject = email.subject

    //             objEmail.HTMLBody = email.body

    //             objEmail.SaveAs(`${outputPath}/managers/${manager[0]}.msg`)

    //             emails.push(objEmail)
    //         }

    //         catch (e) {
    //             console.error(e)
    //             generationMessage.set({ message: "An unknown error occured while generating emails" })
    //         }
    //     })

    //     managerEmails.set(emails)

    //     generationMessage.set({ message: 'Manager emails succesfully generated at ', path: '/output/managers/' })
    // }

    generatingEmails.set(false);
  }, 1000);
};

const getEmployeeEmail = (templateName, keys) => {
  let outlook = new ActiveXObject("Outlook.Application");

  let mailItem = outlook.createItemFromTemplate(templatesDirectory + "\\employee\\" + templateName);

  // let body = mailItem.HTMLBody;

  // Object.keys(keys).forEach((key) => {
  //   mailItem.HTMLBody = body.replaceAll(`{${key}}`, keys[key]);
  // });

  return mailItem;
};

const getManagerEmail = (templateName, { firstName, lastName, impactedEmployees }) => {
  const formattedEmployees = impactedEmployees.map((employee) => {
    return `${employee}`;
  });

  let outlook = new ActiveXObject("Outlook.Application");

  let mailItem = outlook.createItemFromTemplate(templatesDirectory + "\\manager\\" + templateName);

  const keys = {
    firstName,
    lastName,
    impactedEmployees: `${formattedEmployees.join(", ")}`
  };

  let body = mailItem.HTMLBody;

  Object.keys(keys).forEach((key) => {
    body = body.replaceAll(`{${key}}`, keys[key]);
  });

  return { body, subject: mailItem.Subject };
};

const styles = {
  email: `font-size:12pt; font-family:Arial;`
};
