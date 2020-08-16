# google-sheets-utils

Status: Work In Progress

## Description

Utils for Google Sheets to get rid of boilerplate code duplicates in projects.
Implementation of routine tasks migrating from one project to another. 

## Features

* Google Auth with default project credentials.
* Save multiple rows to the sheet.
* Save one row to the sheet.
* Get first sheet ID.
* Clear entire sheet cells.
* New functions will be added as soon as new requirements arise.

##### Additional

* NodeJS
* TypeScript
* Google Sheets API
* [xojs/xo](https://github.com/xojs/xo) with plugins for TypeScript - linting in CLI
* [ESLint](https://github.com/eslint/eslint) - linting in the WebStorm with [ESLint plugin](https://plugins.jetbrains.com/plugin/7494-eslint)

## Installation

```
npm install --save google-sheets-utils
```

## Usage

1. Add your Google Service account to the Google Sheet file editors role.

2. Enable Google Sheets API in your project.

3. Add file `default-credentials.json` with google cloud service account auth key to the root of project. 
This file could be downloaded from Google Cloud IAM console or Firebase Console. 

4. Add `.env` file with contents:
    ```
    GOOGLE_APPLICATION_CREDENTIALS=default-credentials.json
    ```

5. Add `env-cmd` package to the project
    ```
    npm install --save-dev env-cmd
    ```

6. Add `google-sheets-utils` to the project
    ```
    npm install --save google-sheets-utils
    ```

7. Run project with `env-cmd`. Example for cloud functions below:
    ```
    env-cmd npx @google-cloud/functions-framework --target=index --function-signature=myFunction
    ```

##### Code Example:
```typescript
import {GoogleSheetsUtils} from 'google-sheets-utils';

const utils = await GoogleSheetsUtils.create(); // or getInstance() if you want to use a singleton
await utils.clearFirstSheet('yourGoogleSheetsFileId');
await utils.saveRowsToSheet('yourGoogleSheetsFileId', [
  ['A1 cell value', 'B1', 'C1'],
  ['A2', 'B2', 'C2']
]);
```

## License

MIT (c) 2020 Viacheslav Dobromyslov <<viacheslav@dobromyslov.ru>>
