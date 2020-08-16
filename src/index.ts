import {google, sheets_v4} from 'googleapis';
import Sheets = sheets_v4.Sheets;
import {GoogleAuth, GoogleAuthOptions, OAuth2Client} from 'google-auth-library';

export class GoogleSheetsUtils {
  /**
   * Singleton instance.
   */
  protected static instance?: GoogleSheetsUtils;

  /**
   * Google Sheets API instance.
   */
  protected api: Sheets;

  /**
   * Instantiates Google Sheets API with authentication.
   * The constructor is protected.
   * Use GoogleSheetsUtils.create() async static method to authenticate and instantiate new instance.
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
  public static async create(googleAuthOptions: GoogleAuthOptions = {
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  }): Promise<GoogleSheetsUtils> {
    const auth = await google.auth.getClient(googleAuthOptions);
    return new GoogleSheetsUtils(auth);
  }

  /**
   * Creates or gets singleton instance.
   * Use GoogleSheetsUtils.create() if you need new instance.
   * @param googleAuthOptions [OPTIONAL] Authentication options.
   *                                     Default auth scope is https://www.googleapis.com/auth/spreadsheets
   */
  public static async getInstance(googleAuthOptions: GoogleAuthOptions = {
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  }): Promise<GoogleSheetsUtils> {
    if (!GoogleSheetsUtils.instance) {
      GoogleSheetsUtils.instance = await GoogleSheetsUtils.create(googleAuthOptions);
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
   */
  public async clearFirstSheet(fileId: string): Promise<void> {
    await this.api.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      requestBody: {
        requests: [{
          updateCells: {
            range: {
              sheetId: await this.getFirstSheetId(fileId)
            },
            fields: '*'
          }
        }]
      }
    });
  }
}
