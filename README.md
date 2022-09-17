# Spreadsheet builder library

The purpose of this library is to create spreadsheet files using JavaScript/TypeScript code.

```bash
yarn add spreadsheet-builder-lib
```

```bash
npm i spreadsheet-builder-lib
```

Example code:

```typescript
import { buildSpreadsheet } from "spreadsheet-builder-lib";

const spreadsheet = [
  ["String", "Float", "Date", "Time", "Currency", "Percentage"],
  [
    {
      value: "ABBA",
      valueType: "string",
    },
    {
      value: "42.3324",
      valueType: "float",
    },
    {
      value: "2022-02-02",
      valueType: "date",
    },
    {
      value: "19:03:00",
      valueType: "time",
    },
    {
      value: 2.22,
      valueType: "currency",
      currency: "EUR",
    },
    {
      value: 0.4223,
      valueType: "percentage",
    },
  ],
];

const mySpreadsheet = await buildSpreadsheet(spreadsheet);
await writeFile("mySpreadsheet.fods", mySpreadsheet);
```
