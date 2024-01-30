/**
 * @class FileSystem
 */
class FileSystem {
  /**
   * @function loadFile
   * @memberof FileSystem
   *
   * @description Loads a file from the path on disk and returns a binary stream readable by .js Blob handlers.
   *
   * @param {string} path - Path to file
   *
   * @returns {string} A binary string of the file.
   */
  static loadFile = (path: string): string => {
    const stream = new ActiveXObject("ADODB.Stream");

    stream.Open();
    stream.type = 1;
    stream.LoadFromFile(path);

    const binary = stream.Read();

    stream.Close();

    const byteArray = new VBArray(binary).toArray();
    const arrayBuffer = Uint8Array.from(byteArray);
    const buff = arrayBufferToString(arrayBuffer);

    function arrayBufferToString(buffer) {
      const bufferarr = new Uint8Array(buffer);

      const charcode = Array.prototype.slice.apply(bufferarr);

      const binary = charcode.reduce((str, char) => {
        return (str += String.fromCharCode(char));
      }, "");

      return binaryToString(binary);
    }

    function binaryToString(binary) {
      let error;

      try {
        return decodeURIComponent(escape(binary));
      } catch (_error) {
        error = _error;
        if (error instanceof URIError) {
          return binary;
        } else {
          throw error;
        }
      }
    }

    return buff;
  };

  /**
   * @function saveFile
   * @memberof FileSystem
   *
   * @description Saves a Blob to a file with provided extension at the designated path.
   *
   * @param {Blob} blob - Blob to be saved
   * @param {string} path - Where the file should be saved on disk
   * @param {boolean} overwriteFile - Determines if new file should overwrite existing file with same name
   */
  static saveFile = async (blob: Blob, path: string, overwriteFile: boolean): Promise<any> => {
    const buffer = await blobToBuffer(blob);
    const bufferBytes = new Uint8Array(<ArrayBuffer>buffer);
    const overwrite = overwriteFile ? 2 : 1; // Stream Overwrite enum; 1 for no overwrite, 2 for overwrite
    const memory = new ActiveXObject("System.IO.MemoryStream");
    const stream = new ActiveXObject("ADODB.Stream");

    memory.setLength(0);

    return new Promise((resolve) => {
      try {
        bufferBytes.forEach((byte) => memory.writeByte(byte));
        const byteArray = memory.ToArray();

        stream.Open();
        stream.type = 1;

        stream.Write(byteArray);
        stream.SaveToFile(path, overwrite);

        stream.Close();

        resolve(true);
      } catch (e) {
        console.error("Could not save File to " + path);
        console.error(e);

        // @ts-ignore
        reject(false);
      }
    });
  };

  /**
   * @function deleteFile
   * @memberof FileSystem
   *
   * @description Removes file at param filePath if it exists in folder.
   *
   * @param {string} filePath - Where the file should be saved on disk.
   * @return {boolean} true if file was successfully deleted.
   */
  static deleteFile = (filePath: string): boolean => {
    const fso = new ActiveXObject("Scripting.FileSystemObject");
    let fileDeleted: boolean;

    try {
      fso.DeleteFile(filePath);

      fileDeleted = true;
    } catch (e) {
      console.error(e);

      fileDeleted = false;
    }

    return fileDeleted;
  };

  /**
   * @function fileExists
   * @memberof FileSystem
   *
   * @description Checks if a file exists at specified path.
   *
   * @param {string} filePath - file to check for.
   * @return {boolean} true if file exists.
   */
  static fileExists = (filePath: string): boolean => {
    const fso = new ActiveXObject("Scripting.FileSystemObject");

    return fso.FileExists(filePath);
  };

  /**
   * @function folderExists
   * @memberof FileSystem
   *
   * @description Checks if a folder exists at specified path.
   *
   * @param {string} folderPath - folder to check for.
   * @return {boolean} true if folder exists.
   */
  static folderExists = (folderPath: string): boolean => {
    const fso = new ActiveXObject("Scripting.FileSystemObject");

    return fso.folderExists(folderPath);
  };

  /**
   * @function deleteFolder
   * @memberof FileSystem
   *
   * @description Deletes a folder at specified path.
   *
   * @param {string} folderPath - file to check for.
   * @return {boolean} true if folder deleted.
   */
  static deleteFolder = (folderPath: string): boolean => {
    const fso = new ActiveXObject("Scripting.FileSystemObject");
    let folderDeleted = false;

    try {
      fso.DeleteFolder(folderPath);
      folderDeleted = true;
    } catch (e) {
      console.error(e);
    }

    return folderDeleted;
  };

  /**
   * @function run
   * @memberof FileSystem
   *
   * @description Attempts to run a file at the given path.
   *
   * @param {string} filePath  - file to check for.
   * @returns {boolean} true if successfully ran file at path.
   */
  static run = (filePath: string): boolean => {
    const shell = new ActiveXObject("WScript.Shell");
    let runExecuted: boolean;

    try {
      shell.Run(`"${filePath}"`);

      runExecuted = true;
    } catch (e) {
      console.error(e);
      console.error("Unable to run file at path: " + filePath);

      runExecuted = false;
    }

    return runExecuted;
  };

  /**
   * @function createFolder
   * @memberof FileSystem
   *
   * @description Create a folder at the specified folder path
   *
   * @param {string} folderPath - file to check for.
   * @param {boolean} overwrite - Specify whether to overwrite a folder at the given path if it exists
   * @returns {boolean} true if folder successfully created.
   */
  static createFolder = (folderPath: string, overwrite = true): boolean => {
    const fso = new ActiveXObject("Scripting.FileSystemObject");
    let createdFolder: boolean;

    try {
      if (fso.FolderExists(folderPath)) {
        if (overwrite) {
          fso.DeleteFolder(folderPath);

          fso.CreateFolder(folderPath);

          createdFolder = true;
        }
      } else {
        fso.CreateFolder(folderPath);

        createdFolder = true;
      }
    } catch (e) {
      console.error(e);
      console.error("Unable to create folder at path: " + folderPath);

      createdFolder = false;
    }

    return createdFolder;
  };
}

/**
 * Converts a Blob into an ArrayBuffer
 * @param {Blob} blob - Blob to be converted
 *
 * @returns A promise with the resolved blob as an ArrayBuffer
 */
const blobToBuffer = (blob: Blob): Promise<string | ArrayBuffer> => {
  const fileReader: FileReader = new FileReader();

  return new Promise((resolve) => {
    fileReader.onload = (event) => {
      const arrayBuffer = event.target.result;

      resolve(arrayBuffer);
    };
    fileReader.onerror = () => {
      console.error("Error converting blob to buffer");

      // @ts-ignore
      reject("Error converting blob to buffer");
    };

    fileReader.readAsArrayBuffer(blob);
  });
};

export default FileSystem;
