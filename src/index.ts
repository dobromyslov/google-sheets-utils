import {google, sheets_v4} from 'googleapis';
import Sheets = sheets_v4.Sheets;
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
  protected api: Sheets;

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
