export type spreadsheetInput = row[];
export type row = cell[];
export interface cell {
  value: string; // | number
  valueType?: valueType | undefined;
}
export type valueType = "string" | "float" | "date" | "time" | "currency" | "percentage";
export type spreadsheetOutput = string;

export async function buildSpreadsheet(spreadsheet: spreadsheetInput): Promise<string> {
  let s = "";
  spreadsheet.forEach((row) => {
    s += '<table:table-row table:style-name="ro1">';
    row.forEach((cell) => {
      s += tableCell(cell);
    });
    s += "</table:table-row>\n";
  });
  return s;
}

function tableCell(cell: cell): string {
  return `<table:table-cell office:value="${cell.value}" office:value-type="${typeof cell.valueType !== "undefined" ? cell.valueType : "string"}" calcext:value-type="${
    typeof cell.valueType !== "undefined" ? cell.valueType : "string"
  }">
  <text:p>${cell.value}</text:p>
</table:table-cell>
`;
}
