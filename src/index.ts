import {google, sheets_v4} from 'googleapis';
import Sheets = sheets_v4.Sheets;

export class GoogleSheetsUtils {
  /**
   * Google Sheets API instance.
   */
  protected api?: Sheets;

  /**
   * Authorizes, creates, caches and returns Google Sheets API instance.
   */
  public async getApi(): Promise<Sheets> {
    if (!this.api) {
      const auth = await google.auth.getClient({
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      this.api = google.sheets({version: 'v4', auth});
    }

    return this.api;
  }

  /**
   * Saves rows to the sheet.
   * @param fileId Google Sheets file ID.
   * @param rows rows is an array of arrays of values.
   * @param range range name to start from. Default = Sheet1!A1.
   */
  public async saveRowsToSheet(fileId: string, rows: unknown[][], range = 'Sheet1!A1'): Promise<void> {
    const api = await this.getApi();
    await api.spreadsheets.values.update({
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
    const api = await this.getApi();
    const existingSheets = (await api.spreadsheets.get({
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
    const api = await this.getApi();
    await api.spreadsheets.batchUpdate({
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
