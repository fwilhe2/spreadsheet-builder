import { expect, test } from "@jest/globals";
import { exec } from "child_process";
import { readFile, unlink, writeFile } from "fs/promises";
import { promisify } from "util";
import { buildSpreadsheet } from "../lib";

describe("Spreadsheet builder", () => {
  afterEach(async () => {
    await unlink("__tests__/output.csv");
    await unlink("__tests__/output.fods");
  });

  test("Creating a spreadsheet which can be opened using libreoffice - check the csv output is identical", async () => {
    const expectedCsv = `"String","Float","Date","Time","Currency","Currency with Cents","Percentage"\n"ABBA",42.3324,2022-02-02,19:03:00,3.00€,2.22€,42.23%\n`;

    const spreadsheet = JSON.parse((await readFile("__tests__/data-formats.json")).toString());
    const template = (await readFile("empty-file.fods")).toString();
    const spreadsheetRows = await buildSpreadsheet(spreadsheet);
    const output = template.replace("TABLE_ROWS", spreadsheetRows);
    await writeFile("__tests__/output.fods", output);

    // todo: see why this did not work using execa

    const e = promisify(exec);
    await e('libreoffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":"44,34,76,1,,1031,true,true" __tests__/output.fods --outdir __tests__');

    const actualCsv = (await readFile("__tests__/output.csv")).toString();
    expect(actualCsv).toEqual(expectedCsv);
  });
});
