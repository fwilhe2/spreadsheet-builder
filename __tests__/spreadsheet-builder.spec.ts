import { expect, test } from "@jest/globals";
import { exec } from "child_process";
import { mkdir, readFile, rm, writeFile } from "fs/promises";
import { promisify } from "util";
import { buildSpreadsheet, spreadsheetInput } from "../lib";

describe("Spreadsheet builder", () => {
  beforeAll(async () => {
    await rm("__tests__/output", { recursive: true, force: true });
    await mkdir("__tests__/output");
  });

  test("Creating a spreadsheet which can be opened using libreoffice - check the csv output is identical", async () => {
    const expectedCsv = `"String","Float","Date","Time","Currency","Currency with Cents","Percentage"\n"ABBA",42.33,2022-02-02,19:03:00,3.00€,2.22€,42.23%\n`;

    const spreadsheet = JSON.parse((await readFile("__tests__/data-formats.json")).toString());
    const actualFods = await buildSpreadsheet(spreadsheet);
    await writeFile("__tests__/output/common-data-formats.fods", actualFods);

    // todo: see why this did not work using execa

    const e = promisify(exec);
    await e('libreoffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":"44,34,76,1,,1031,true,true" __tests__/output/common-data-formats.fods --outdir __tests__/output');

    const actualCsv = (await readFile("__tests__/output/common-data-formats.csv")).toString();
    expect(actualCsv).toEqual(expectedCsv);
  });

  test("Performance Model Spreadsheet", async () => {
    const expectedCsv = `"Number of CPUs","Parallel Computing Time","Sequential Computing Time","Speedup","Efficiency"\n4.00,"25,800.00","100,000.00",3.88,0.97\n5.00,"21,000.00","100,000.00",4.76,0.95\n6.00,"17,866.67","100,000.00",5.60,0.93\n`;
    const problemSizeX = 100;
    const problemSizeY = 100;
    const calculationTimePerCell = 10;
    const numberOfOperations = 1;
    const communicationTimePerCell = 200;
    const mySpreadsheet: spreadsheetInput = [["Number of CPUs", "Parallel Computing Time", "Sequential Computing Time", "Speedup", "Efficiency"]];
    for (let numberOfCpus = 4; numberOfCpus < 7; numberOfCpus++) {
      const timeParallel = (problemSizeX / numberOfCpus) * problemSizeY * calculationTimePerCell * numberOfOperations + communicationTimePerCell * numberOfCpus;
      const timeSequential = problemSizeX * problemSizeY * calculationTimePerCell * numberOfOperations;
      const speedup = timeSequential / timeParallel;
      mySpreadsheet.push([
        {
          value: numberOfCpus.toString(),
          valueType: "float",
        },
        {
          value: `${timeParallel}`,
          valueType: "float",
        },
        {
          value: `${timeSequential}`,
          valueType: "float",
        },
        {
          value: `${speedup}`,
          valueType: "float",
        },
        {
          value: `${speedup / numberOfCpus}`,
          valueType: "float",
        },
      ]);
    }
    const actualFods = await buildSpreadsheet(mySpreadsheet);
    await writeFile("__tests__/output/performanceModel.fods", actualFods);

    // todo: see why this did not work using execa

    const e = promisify(exec);
    await e('libreoffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":"44,34,76,1,,1031,true,true" __tests__/output/performanceModel.fods --outdir __tests__/output');

    const actualCsv = (await readFile("__tests__/output/performanceModel.csv")).toString();
    expect(actualCsv).toEqual(expectedCsv);
  });

  test("formula", async () => {
    const mySpreadsheet: spreadsheetInput = [
      ['sum', 'avg'],
      [{ "value": "42.3324", "valueType": "float" }, { "value": "42.3324", "valueType": "float" }],
      [{ "value": "1.12", "valueType": "float" }, { "value": "1.12", "valueType": "float" }],
      [{ "value": "7.98", "valueType": "float" }, { "value": "7.98", "valueType": "float" }],
      [
        {
          "formula": "=SUM([.A2:.A4])",
          "valueType": "float"
        },
        {
          "formula": "=AVERAGE([.B2:.B4])",
          "valueType": "float"
        }
      ],
      [
        {
          "formula": "=[.A5]+1",
          "valueType": "float"
        },
        {
          "formula": "=ROUND([.A6];1)",
          "valueType": "float"
        }
      ],

    ]
    const actualFods = await buildSpreadsheet(mySpreadsheet);
    await writeFile("__tests__/output/formula.fods", actualFods);

  })
});
