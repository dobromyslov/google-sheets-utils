import {google, sheets_v4} from 'googleapis';
import {GoogleAuth, GoogleAuthOptions, OAuth2Client} from 'google-auth-library';

export class GoogleSheetsUtils {
  /**
   * Singleton instance.
   */
  protected static instance?: GoogleSheetsUtils;

  /**
   * Authentication options used on first singleton initialization.
   */
  protected static singletonGoogleAuthOptions?: GoogleAuthOptions;

  /**
   * Google Sheets API instance.
   */
  protected api: sheets_v4.Sheets;

  /**
   * Instantiates Google Sheets API with authentication.
   * The constructor is protected.
   * Use GoogleSheetsUtils.create() async static method to authenticate and create new instance
   * or GoogleSheetsUtils.getInstance() if you need a singleton.
   * @param auth authentication object.
   */
  protected constructor(auth: GoogleAuth | OAuth2Client | string) {
    this.api = google.sheets({version: 'v4', auth});
  }

  /**
   * Authenticates and creates new instance.
   * Use GoogleSheetsUtils.getInstance() if you need a singleton instance.
   * @param googleAuthOptions [OPTIONAL] Authentication options.
   *                                     Default auth scope is https://www.googleapis.com/auth/spreadsheets
   */
  public static async create(googleAuthOptions?: GoogleAuthOptions): Promise<GoogleSheetsUtils> {
    // Set default auth scope
    googleAuthOptions = googleAuthOptions ?? {};
    googleAuthOptions.scopes = googleAuthOptions.scopes ?? ['https://www.googleapis.com/auth/spreadsheets'];

    const auth = await google.auth.getClient(googleAuthOptions);
    return new GoogleSheetsUtils(auth);
  }

  /**
   * Creates or gets singleton instance.
   * Use GoogleSheetsUtils.create() if you need new instance.
   * @param googleAuthOptions [OPTIONAL] Authentication options.
   *                                     Default auth scope is https://www.googleapis.com/auth/spreadsheets
   */
  public static async getInstance(googleAuthOptions?: GoogleAuthOptions): Promise<GoogleSheetsUtils> {
    if (!GoogleSheetsUtils.instance) {
      GoogleSheetsUtils.instance = await GoogleSheetsUtils.create(googleAuthOptions);
      GoogleSheetsUtils.singletonGoogleAuthOptions = googleAuthOptions;
    } else if (googleAuthOptions && JSON.stringify(googleAuthOptions) !== JSON.stringify(GoogleSheetsUtils.singletonGoogleAuthOptions)) {
      throw new Error(
        'Singleton instance has been already created with authOptions. ' +
        'Please use GoogleSheetsUtils.create() to create another instance with new different auth options.'
      );
    }

    return GoogleSheetsUtils.instance;
  }

  /**
   * Returns Google Sheets raw API.
   */
  public getApi(): sheets_v4.Sheets {
    return this.api;
  }

  /**
   * Returns values from the specified spreadsheet and range.
   * @param fileId
   * @param range [OPTIONAL] range in A1 notation. Default is A1.
   */
  public async getRowsFromSheet(fileId: string, range = 'A1'): Promise<unknown[][]> {
    return (await this.api.spreadsheets.values.get({
      spreadsheetId: fileId,
      range
    })).data.values ?? [];
  }

  /**
   * Saves rows to the sheet.
   * @param fileId Google Sheets file ID.
   * @param rows rows is an array of arrays of values.
   * @param range range name to start from. Default = Sheet1!A1.
   */
  public async saveRowsToSheet(fileId: string, rows: unknown[][], range = 'Sheet1!A1'): Promise<void> {
    await this.api.spreadsheets.values.update({
      spreadsheetId: fileId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: rows
      }
    });
  }

  /**
   * Saves one row values to the sheet.
   * @param fileId Google Sheets file ID.
   * @param values array of values (for columns of one row).
   * @param range row cell address to update values starting from. Default = Sheet1!A1
   */
  public async saveRowToSheet(fileId: string, values: unknown[], range = 'Sheet1!A1'): Promise<void> {
    await this.saveRowsToSheet(fileId, [values], range);
  }

  /**
   * Returns first sheet ID.
   * @param fileId Google Sheets file ID.
   */
  public async getFirstSheetId(fileId: string): Promise<number | null> {
    const existingSheets = (await this.api.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties'
    }))?.data?.sheets;

    if (existingSheets && existingSheets.length > 0) {
      for (const sheet of existingSheets) {
        if (sheet.properties?.index === 0) {
          return sheet.properties?.sheetId ?? null;
        }
      }
    }

    return null;
  }

  /**
   * Returns sheet ID by it's title.
   * @param fileId Google Sheets file ID.
   * @param title Title of the sheet.
   */
  public async getSheetIdByTitle(fileId: string, title: string): Promise<number | null> {
    const existingSheets = (await this.api.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties'
    }))?.data?.sheets;

    if (existingSheets && existingSheets.length > 0) {
      for (const sheet of existingSheets) {
        if (sheet.properties?.title === title) {
          return sheet.properties?.sheetId ?? null;
        }
      }
    }

    return null;
  }

  /**
   * Copies sheet to the specified file.
   * @param fileId Source Google Sheets file ID.
   * @param sheetId Source sheet ID.
   * @param destinationFileId Destination Google Sheets file ID.
   */
  public async copySheet(fileId: string, sheetId: number, destinationFileId: string): Promise<void> {
    await this.api.spreadsheets.sheets.copyTo({
      spreadsheetId: fileId,
      sheetId,
      requestBody: {
        destinationSpreadsheetId: destinationFileId
      }
    });
  }

  /**
   * Copies sheet with the given title to the destination file.
   * @param fileId Source Google Sheets file ID.
   * @param sheetTitle Source sheet title.
   * @param destinationFileId Destination Google Sheets file ID.
   */
  public async copySheetByTitle(fileId: string, sheetTitle: string, destinationFileId: string): Promise<void> {
    const sheetId = await this.getSheetIdByTitle(fileId, sheetTitle);
    if (sheetId === null) {
      throw new Error(`Sheet '${sheetTitle}' not found.`);
    }

    await this.copySheet(fileId, sheetId, destinationFileId);
  }

  /**
   * Finds sheet with the given ID and renames it.
   * @param fileId Google Sheets file ID.
   * @param sheetId Sheet ID.
   * @param newTitle New sheet title.
   */
  public async renameSheetById(fileId: string, sheetId: number, newTitle: string): Promise<void> {
    await this.api.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      requestBody: {
        requests: [{
          updateSheetProperties: {
            properties: {
              sheetId,
              title: newTitle
            }
          }
        }]
      }
    });
  }

  /**
   * Finds sheet with the given title and renames it.
   * @param fileId Google Sheets file ID.
   * @param currentTitle Current title to find sheet.
   * @param newTitle New sheet title.
   */
  public async renameSheetByTitle(fileId: string, currentTitle: string, newTitle: string): Promise<void> {
    const sheetId = await this.getSheetIdByTitle(fileId, currentTitle);
    if (sheetId === null) {
      throw new Error(`Sheet '${currentTitle}' not found.`);
    }

    await this.renameSheetById(fileId, sheetId, newTitle);
  }

  /**
   * Deletes sheet with the given ID.
   * @param fileId Google Sheets file ID.
   * @param sheetId Sheet id to delete.
   */
  public async deleteSheetById(fileId: string, sheetId: number): Promise<void> {
    await this.api.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      requestBody: {
        requests: [{
          deleteSheet: {
            sheetId
          }
        }]
      }
    });
  }

  /**
   * Finds sheet by the title and deletes it.
   * @param fileId Google Sheets file ID.
   * @param title Title to find the sheet.
   */
  public async deleteSheetByTitle(fileId: string, title: string): Promise<void> {
    const sheetId = await this.getSheetIdByTitle(fileId, title);
    if (sheetId === null) {
      throw new Error(`Sheet '${title}' not found.`);
    }

    await this.deleteSheetById(fileId, sheetId);
  }

  /**
   * Clears data on the first sheet.
   * @param fileId Google Sheets file ID.
   * @param startRowIndex start from row. Default = 1.
   * @param startColumnIndex start from column. Default = 1.
   */
  public async clearFirstSheet(fileId: string, startRowIndex = 1, startColumnIndex = 1): Promise<void> {
    await this.api.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      requestBody: {
        requests: [{
          updateCells: {
            range: {
              sheetId: await this.getFirstSheetId(fileId),
              startRowIndex,
              startColumnIndex
            },
            // See: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells
            fields: 'userEnteredValue'
          }
        }]
      }
    });
  }

  /**
   * Updates rows height.
   * @param fileId spreadsheet file ID.
   * @param height row height in pixels.
   * @param startRowIndex starting from row index. Default = 1.
   */
  public async setRowsHeight(fileId: string, height: number, startRowIndex = 1): Promise<void> {
    await this.api.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      requestBody: {
        requests: [{
          updateDimensionProperties: {
            range: {
              sheetId: await this.getFirstSheetId(fileId),
              dimension: 'ROWS',
              startIndex: startRowIndex
            },
            properties: {
              pixelSize: height
            },
            fields: 'pixelSize'
          }
        }]
      }
    });
  }
}
