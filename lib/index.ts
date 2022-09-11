export type spreadsheetInput = row[];
export type row = cell[];
export type cell = complexCell | string;
export interface complexCell {
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
  return `<table:table-cell ${tableCellAttributes(cell)}>
  <text:p>${cellValue(cell)}</text:p>
</table:table-cell>
`;
}

function tableCellAttributes(cell: cell) {
  return `${cellStyle(cell)} ${cellValueAttribute(cell)} ${dateValue(cell)} ${timeValue(cell)} ${currencyValue(cell)} office:value-type="${cellValueType(cell)}" calcext:value-type="${cellValueType(
    cell
  )}"`;
}

function cellStyle(cell: cell): string {
  if (typeof cell == "string") {
    return "";
  }

  if (cell.valueType === "date") {
    return 'table:style-name="ce1"';
  }

  if (cell.valueType === "time") {
    return 'table:style-name="ce2"';
  }

  if (cell.valueType === "currency") {
    return 'table:style-name="ce3"';
  }

  if (cell.valueType === "percentage") {
    return 'table:style-name="ce4"';
  }

  return "";
}

function dateValue(cell: cell): string {
  if (typeof cell === "string" || cell.valueType !== "date") {
    return "";
  }

  return `office:date-value="${cell.value}"`;
}

function timeValue(cell: cell): string {
  if (typeof cell === "string" || cell.valueType !== "time") {
    return "";
  }

  // assume hh:mm:ss format for now
  const components = cell.value.split(":");
  if (components.length != 3) {
    console.warn("expected hh:mm:ss format");
  }

  return `office:time-value="PT${components[0]}H${components[1]}M${components[2]}S"`;
}

function currencyValue(cell: cell): string {
  if (typeof cell === "string" || cell.valueType !== "currency") {
    return "";
  }

  return `office:currency="EUR"`;
}

function cellValue(cell: cell): string {
  if (typeof cell == "string") {
    return cell;
  }

  return cell.value;
}

function cellValueAttribute(cell: cell): string {
  if (typeof cell == "string") {
    return `office:value="${cell}"`;
  }

  if (cell.valueType === "date" || cell.valueType === "time") {
    return "";
  }

  return `office:value="${cell.value}"`;
}

function cellValueType(cell: cell): valueType {
  if (typeof cell == "string" || cell.valueType === undefined) {
    return "string";
  }

  return cell.valueType;
}
