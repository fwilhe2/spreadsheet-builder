export type spreadsheetInput = row[];
export type row = cell[];
export type cell = complexCell | string;
export interface complexCell {
  value: string; // | number
  valueType?: valueType | undefined;
}
export type valueType = "string" | "float" | "date" | "time" | "currency" | "percentage";
export type spreadsheetOutput = string;

/**
 * Build a spreadsheet from data
 * @param spreadsheet list of lists of cells
 * @returns string Flat OpenDocument Spreadsheet document
 */
export async function buildSpreadsheet(spreadsheet: spreadsheetInput): Promise<string> {
  let tableRows = "";
  spreadsheet.forEach((row) => {
    tableRows += '<table:table-row table:style-name="ro1">';
    row.forEach((cell) => {
      tableRows += tableCell(cell);
    });
    tableRows += "</table:table-row>\n";
  });
  return FODS_TEMPLATE.replace("TABLE_ROWS", tableRows);
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

const FODS_TEMPLATE = `<?xml version="1.0" encoding="UTF-8"?>

<office:document xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:rpt="http://openoffice.org/2005/report" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xforms="http://www.w3.org/2002/xforms" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
    <office:meta>
        <meta:generator>My Spreadsheet Builder</meta:generator>
    </office:meta>


    <office:styles>
        <style:default-style style:family="table-cell">
            <style:paragraph-properties style:tab-stop-distance="0.5in" />
            <style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="en" fo:country="DE" style:font-name-asian="Noto Sans CJK SC" style:font-size-asian="10pt" style:language-asian="zh" style:country-asian="CN" style:font-name-complex="Lohit Devanagari" style:font-size-complex="10pt" style:language-complex="hi" style:country-complex="IN" />
        </style:default-style>
        <number:number-style style:name="N0">
            <number:number number:min-integer-digits="1" />
        </number:number-style>
        <style:style style:name="Default" style:family="table-cell" />
        <style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default">
            <style:text-properties fo:color="#000000" fo:font-size="24pt" fo:font-style="normal" fo:font-weight="bold" style:font-size-asian="24pt" style:font-style-asian="normal" style:font-weight-asian="bold" style:font-size-complex="24pt" style:font-style-complex="normal" style:font-weight-complex="bold" />
        </style:style>
        <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="table-cell" style:parent-style-name="Heading">
            <style:text-properties fo:font-size="18pt" style:font-size-asian="18pt" style:font-size-complex="18pt" />
        </style:style>
        <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="table-cell" style:parent-style-name="Heading">
            <style:text-properties fo:font-size="12pt" style:font-size-asian="12pt" style:font-size-complex="12pt" />
        </style:style>
        <style:style style:name="Text" style:family="table-cell" style:parent-style-name="Default" />
        <style:style style:name="Note" style:family="table-cell" style:parent-style-name="Text">
            <style:table-cell-properties fo:background-color="#ffffcc" style:diagonal-bl-tr="none" style:diagonal-tl-br="none" fo:border="0.74pt solid #808080" />
            <style:text-properties fo:color="#333333" />
        </style:style>
        <style:style style:name="Footnote" style:family="table-cell" style:parent-style-name="Text">
            <style:text-properties fo:color="#808080" fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic" />
        </style:style>
        <style:style style:name="Hyperlink" style:family="table-cell" style:parent-style-name="Text">
            <style:text-properties fo:color="#0000ee" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="#0000ee" />
        </style:style>
        <style:style style:name="Status" style:family="table-cell" style:parent-style-name="Default" />
        <style:style style:name="Good" style:family="table-cell" style:parent-style-name="Status">
            <style:table-cell-properties fo:background-color="#ccffcc" />
            <style:text-properties fo:color="#006600" />
        </style:style>
        <style:style style:name="Neutral" style:family="table-cell" style:parent-style-name="Status">
            <style:table-cell-properties fo:background-color="#ffffcc" />
            <style:text-properties fo:color="#996600" />
        </style:style>
        <style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Status">
            <style:table-cell-properties fo:background-color="#ffcccc" />
            <style:text-properties fo:color="#cc0000" />
        </style:style>
        <style:style style:name="Warning" style:family="table-cell" style:parent-style-name="Status">
            <style:text-properties fo:color="#cc0000" />
        </style:style>
        <style:style style:name="Error" style:family="table-cell" style:parent-style-name="Status">
            <style:table-cell-properties fo:background-color="#cc0000" />
            <style:text-properties fo:color="#ffffff" fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold" />
        </style:style>
        <style:style style:name="Accent" style:family="table-cell" style:parent-style-name="Default">
            <style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold" />
        </style:style>
        <style:style style:name="Accent_20_1" style:display-name="Accent 1" style:family="table-cell" style:parent-style-name="Accent">
            <style:table-cell-properties fo:background-color="#000000" />
            <style:text-properties fo:color="#ffffff" />
        </style:style>
        <style:style style:name="Accent_20_2" style:display-name="Accent 2" style:family="table-cell" style:parent-style-name="Accent">
            <style:table-cell-properties fo:background-color="#808080" />
            <style:text-properties fo:color="#ffffff" />
        </style:style>
        <style:style style:name="Accent_20_3" style:display-name="Accent 3" style:family="table-cell" style:parent-style-name="Accent">
            <style:table-cell-properties fo:background-color="#dddddd" />
        </style:style>
        <style:style style:name="Result" style:family="table-cell" style:parent-style-name="Default">
            <style:text-properties fo:font-style="italic" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-weight="bold" style:font-style-asian="italic" style:font-weight-asian="bold" style:font-style-complex="italic" style:font-weight-complex="bold" />
        </style:style>
    </office:styles>
    <office:automatic-styles>
        <style:style style:name="co1" style:family="table-column">
            <style:table-column-properties fo:break-before="auto" style:column-width="0.889in" />
        </style:style>
        <style:style style:name="ro1" style:family="table-row">
            <style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true" />
        </style:style>
        <style:style style:name="ta1" style:family="table" style:master-page-name="Default">
            <style:table-properties table:display="true" style:writing-mode="lr-tb" />
        </style:style>
        <number:currency-style style:name="N10121P0" style:volatile="true" number:language="en" number:country="DE">
            <number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true" />
            <number:text />
            <number:currency-symbol number:language="de" number:country="DE">€</number:currency-symbol>
        </number:currency-style>
        <number:currency-style style:name="N10121" number:language="en" number:country="DE">
            <style:text-properties fo:color="#ff0000" />
            <number:text>-</number:text>
            <number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true" />
            <number:text />
            <number:currency-symbol number:language="de" number:country="DE">€</number:currency-symbol>
            <style:map style:condition="value()&gt;=0" style:apply-style-name="N10121P0" />
        </number:currency-style>
        <number:percentage-style style:name="N11">
            <number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" />
            <number:text>%</number:text>
        </number:percentage-style>
        <number:date-style style:name="N49">
            <number:year number:style="long" />
            <number:text>-</number:text>
            <number:month number:style="long" />
            <number:text>-</number:text>
            <number:day number:style="long" />
        </number:date-style>
        <number:time-style style:name="N61">
            <number:hours number:style="long" />
            <number:text>:</number:text>
            <number:minutes number:style="long" />
            <number:text>:</number:text>
            <number:seconds number:style="long" />
        </number:time-style>
        <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N49" />
        <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N61" />
        <style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N10121" />
        <style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N11" />
        <style:page-layout style:name="pm1">
            <style:page-layout-properties style:writing-mode="lr-tb" />
            <style:header-style>
                <style:header-footer-properties fo:min-height="0.2953in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-bottom="0.0984in" />
            </style:header-style>
            <style:footer-style>
                <style:header-footer-properties fo:min-height="0.2953in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0.0984in" />
            </style:footer-style>
        </style:page-layout>
    </office:automatic-styles>

    <office:body>
        <office:spreadsheet>
            <table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true" />
            <table:table table:name="Sheet1" table:style-name="ta1">
                TABLE_ROWS
            </table:table>
            <table:named-expressions />
        </office:spreadsheet>
    </office:body>
</office:document>`;
