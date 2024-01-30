import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import FileSystem from "./filesystem";
import expressionParser from "docxtemplater/expressions-ie11.js";

export enum TemplateError {
    MISSING_LOCKED_FILE,
    INVALID_TEMPLATE,
    UNKNOWN_ERROR,
}

/** 
 * Creates a template from
 * 
 * @param {*} filePath 
 * @param {*} templateData 
 * @returns {Blob} A generated .docx template file as a blob
 */
export const createTemplate = (filePath: string, templateData: Blob): Blob | TemplateError => {
    let doc: Docxtemplater;

    try {
        const file: string = FileSystem.loadFile(filePath);
        const zip: PizZip = new PizZip();

        zip.load(file);

        doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            parser: expressionParser,
        });

        doc.render({ ...templateData });
    } catch (error) {
        console.log(error)

        switch (error.description) {
            case "Unable to get property 'toString' of undefined or null reference":
                return TemplateError.INVALID_TEMPLATE;
            case "File could not be opened.":
                return TemplateError.MISSING_LOCKED_FILE;
            default:
                return TemplateError.UNKNOWN_ERROR;
        }
    }

    const blob: Blob = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        compression: "DEFLATE",
    });

    return blob;
};
