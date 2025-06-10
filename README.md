# gas-budget-scripts

Google App Scripts for budgeting

### Summary

This repo contains a Google App Script TypeScript project that can be used on a Google Sheets Spreadsheet with a [Container-bound app script project](https://developers.google.com/apps-script/guides/bound). Set up the connection between this repo and the app script project with [these instructions](https://developers.google.com/apps-script/guides/clasp).

#### Setup

To install dependencies:

```bash
npm i
```

To log into Google with Google's clasp CLI tool:

```bash
npm run login
```

To connect this project to your App Script project (will create a `.clasp.json` file that has the appropriate App Script project ID):

```bash
clasp clone SCRIPT_PROJECT_ID
```

#### Development

To make changes to your connected App Script project:

```bash
npm run push
```

To watch your files and make changes to your connected App Script project as you change files:

```bash
npm start
```

To open the connected App Script project:

```bash
npm run open
```

See `package.json` for more operations.

### Branches

- `main` branch - current development. Use with a test budget sheet
- `deploy-budget-sheet` - current production code. Use with real budget sheet.
