import { Component } from '@angular/core';
import * as ExcelJS from "exceljs";
import * as FileSaver from 'file-saver';
import { Filesystem, Directory, Encoding } from '@capacitor/filesystem';
import {Capacitor} from "@capacitor/core";
import write_blob from "capacitor-blob-writer";

@Component({
  selector: 'app-tab1',
  templateUrl: 'tab1.page.html',
  styleUrls: ['tab1.page.scss']
})
export class Tab1Page {

 

  constructor() {}

  ngOnInit(){}

    // From the PoC we can conclude this is absolutely the package
// that supports all the required features:
// - merge cells
// - font / fill color changes
// - auto filters
// - formula values
// - conditional formatting
// - number formatting
// The package has an MIT license which makes it suitable for commercial / propiertary software projects.
// Therefore, this is package recommended by the tech team for this project.
// URL: https://www.npmjs.com/package/exceljs

generateExcelWithExcelJs(): void {

   // FOR EXCEL JS COLOR IS REPRESENTED IN ARGB, THIS WEBSITE HAS A VERY HANDY ONLINE CONVERTER:
  // https://www.myfixguide.com/color-converter/
  const BLOL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
  const TITLE_1_FONT = { name: 'Calibri', size: 35, color: { argb: 'FFFFFFFF' }, bold: true };
  const TITLE_2_FONT = { name: 'Calibri', size: 11, color: { argb: 'FFFFFFFF' }, bold: true };
  const COUNT_FONT = { name: 'Calibri', size: 35, bold: true };
  const BLUE_CELL_BACKGROUND : ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B3BC' } };
  const ALIGN_CELL_CENTER_CENTER : Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'center' };
  const THIN_CELL_BORDER : Partial<ExcelJS.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const ONE_DECIMAL_POS_FORMAT = '#,#0.0';

  const columns = [
    { width: 4.83 },
    { width: 10.83 },
    { width: 7.33 }, // in characters
    { width: 4.83 },
    { width: 4.83 },
    { width: 4.83 },
    { width: 6.50 },
    { width: 11.17 },
    { width: 6.83 },
    { width: 8.83 },
    { width: 9.50 },
    { width: 9.50 },
    { width: 9.50 },
    { width: 11.50 }
  ]


  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('6', { properties: { tabColor: { argb: 'FFC0000' } } });

  worksheet.columns = columns;

  // ROW 1
  const row1 = worksheet.getRow(1);
  row1.height = 36;

  const a1 = worksheet.getCell('A1');
  a1.value = 'CHEQUEO DE MASTITIS';
  a1.font = TITLE_1_FONT;
  a1.alignment = ALIGN_CELL_CENTER_CENTER;
  a1.fill = BLUE_CELL_BACKGROUND;

  const n1 = worksheet.getCell('N1');
  n1.value = 6;
  n1.font = TITLE_1_FONT;
  n1.alignment = ALIGN_CELL_CENTER_CENTER;
  n1.fill = BLUE_CELL_BACKGROUND;

  worksheet.mergeCells('A1:M1');

  //row1.commit(); // I have no idea what this line does or why it would be required

  // ROW 1 END

  // ROW 2
  const row2 = worksheet.getRow(2);
  row2.height = 15;

  const a2 = worksheet.getCell('A2');
  a2.value = 'PROPIETARIO:';
  a2.font = TITLE_2_FONT;
  a2.alignment = ALIGN_CELL_CENTER_CENTER;
  a2.fill = BLUE_CELL_BACKGROUND;

  const k2 = worksheet.getCell('K2');
  k2.value = '$/Lt:';
  k2.font = TITLE_2_FONT;
  k2.alignment = ALIGN_CELL_CENTER_CENTER;
  k2.fill = BLUE_CELL_BACKGROUND;

  const m2 = worksheet.getCell('M2');
  m2.value = 'FECHA:';
  m2.font = TITLE_2_FONT;
  m2.alignment = ALIGN_CELL_CENTER_CENTER;
  m2.fill = BLUE_CELL_BACKGROUND;

  worksheet.mergeCells('A2:B2');
  worksheet.mergeCells('C2:J2');

  // ROW 2 END


  // ROW 3
  const a3 = worksheet.getCell('A3');
  a3.value = 'No.';
  a3.font = TITLE_2_FONT;
  a3.alignment = ALIGN_CELL_CENTER_CENTER;
  a3.fill = BLUE_CELL_BACKGROUND;

  const b3 = worksheet.getCell('B3');
  b3.value = 'Identificación';
  b3.font = TITLE_2_FONT;
  b3.alignment = ALIGN_CELL_CENTER_CENTER;
  b3.fill = BLUE_CELL_BACKGROUND;

  const c3 = worksheet.getCell('C3');
  c3.value = 'DI';
  c3.font = TITLE_2_FONT;
  c3.alignment = ALIGN_CELL_CENTER_CENTER;
  c3.fill = BLUE_CELL_BACKGROUND;

  const d3 = worksheet.getCell('D3');
  d3.value = 'DD';
  d3.font = TITLE_2_FONT;
  d3.alignment = ALIGN_CELL_CENTER_CENTER;
  d3.fill = BLUE_CELL_BACKGROUND;

  const e3 = worksheet.getCell('E3');
  e3.value = 'PI';
  e3.font = TITLE_2_FONT;
  e3.alignment = ALIGN_CELL_CENTER_CENTER;
  e3.fill = BLUE_CELL_BACKGROUND;

  const f3 = worksheet.getCell('F3');
  f3.value = 'PD';
  f3.font = TITLE_2_FONT;
  f3.alignment = ALIGN_CELL_CENTER_CENTER;
  f3.fill = BLUE_CELL_BACKGROUND;

  const g3 = worksheet.getCell('G3');
  g3.value = 'PDN';
  g3.font = TITLE_2_FONT;
  g3.alignment = ALIGN_CELL_CENTER_CENTER;
  g3.fill = BLUE_CELL_BACKGROUND;

  const h3 = worksheet.getCell('H3');
  h3.value = 'DIAGNOSTICO';
  h3.font = TITLE_2_FONT;
  h3.alignment = ALIGN_CELL_CENTER_CENTER;
  h3.fill = BLUE_CELL_BACKGROUND;

  const i3 = worksheet.getCell('I3');
  i3.value = 'PARTO';
  i3.font = TITLE_2_FONT;
  i3.alignment = ALIGN_CELL_CENTER_CENTER;
  i3.fill = BLUE_CELL_BACKGROUND;

  const j3 = worksheet.getCell('J3');
  j3.value = 'D.E.L';
  j3.font = TITLE_2_FONT;
  j3.alignment = ALIGN_CELL_CENTER_CENTER;
  j3.fill = BLUE_CELL_BACKGROUND;

  const k3 = worksheet.getCell('K3');
  k3.value = 'APORTE';
  k3.font = TITLE_2_FONT;
  k3.alignment = ALIGN_CELL_CENTER_CENTER;
  k3.fill = BLUE_CELL_BACKGROUND;

  const l3 = worksheet.getCell('L3');
  l3.value = 'ESTADO';
  l3.font = TITLE_2_FONT;
  l3.alignment = ALIGN_CELL_CENTER_CENTER;
  l3.fill = BLUE_CELL_BACKGROUND;

  const m3 = worksheet.getCell('M3');
  m3.value = 'C.E';
  m3.font = TITLE_2_FONT;
  m3.alignment = ALIGN_CELL_CENTER_CENTER;
  m3.fill = BLUE_CELL_BACKGROUND;

  const n3 = worksheet.getCell('N3');
  n3.value = 'HATO';
  n3.font = TITLE_2_FONT;
  n3.alignment = ALIGN_CELL_CENTER_CENTER;
  n3.fill = BLUE_CELL_BACKGROUND;
  // ROW 3 END

  // ROW 4
  const a4 = worksheet.getCell('A4');
  a4.value = 2;
  a4.font = TITLE_2_FONT;
  a4.alignment = ALIGN_CELL_CENTER_CENTER;
  a4.fill = BLUE_CELL_BACKGROUND;

  const b4 = worksheet.getCell('B4');
  b4.value = 953;
  b4.alignment = ALIGN_CELL_CENTER_CENTER;

  const g4 = worksheet.getCell('G4');
  g4.value = 0.0;
  g4.numFmt = ONE_DECIMAL_POS_FORMAT;
  g4.alignment = ALIGN_CELL_CENTER_CENTER;

  const i4 = worksheet.getCell('I4');
  i4.value = 4;
  i4.alignment = ALIGN_CELL_CENTER_CENTER;

  const j4 = worksheet.getCell('J4');
  j4.value = { formula: 'I4 * 6' } as ExcelJS.CellValue;
  j4.alignment = ALIGN_CELL_CENTER_CENTER;

  const m4 = worksheet.getCell('M4');
  m4.value = 6.4;
  m4.numFmt = ONE_DECIMAL_POS_FORMAT;
  m4.alignment = ALIGN_CELL_CENTER_CENTER;

  const n4 = worksheet.getCell('N4');
  n4.value = 1;
  n4.alignment = ALIGN_CELL_CENTER_CENTER;

  // ROW 4 END

  // ROW 7
  const a7 = worksheet.getCell('A7');
  a7.value = 1116;
  a7.font = COUNT_FONT;
  a7.alignment = ALIGN_CELL_CENTER_CENTER;

  worksheet.mergeCells('A7:B10');

  // ROW 7 END

  //ADD THIN BORDERS TO ALL ROWS and CELLS that have a value
  worksheet.eachRow(function (row, rowNumber) {
    row.eachCell(function (cell, colNumber) {
      cell.border = THIN_CELL_BORDER;
    });
  });

  // ADD AUTOFILTERS FOR THE SPECIFIED RANGE
  worksheet.autoFilter = {
    from: 'A3',
    to: 'N3',
  };

  // ADD FREEZE ROWS / COLS
  worksheet.views = [
    { state: 'frozen', xSplit: 14, ySplit: 3 }
  ];

  const conditionalFormattingRule : ExcelJS.ConditionalFormattingRule[] = [
    {
      type: 'expression',
      formulae: ['$I4=4'],
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFEB9C' } }, font: { name: 'Calibri', size: 11, color: { argb: 'FF9C6500' }, bold: false } }
    } as ExcelJS.ConditionalFormattingRule
  ];

  // ADD CONDITIONAL FORMATTING
  worksheet.addConditionalFormatting({
    ref: 'I4:I4',
    rules: conditionalFormattingRule
  });

  /*workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], { type: BLOL_TYPE });
    FileSaver.saveAs(blob, 'exceljs.xlsx');
  }).catch((err) => {
    console.log('Something really bad happened writing the XlSX file.', err);
  });*/


  workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], { type: BLOL_TYPE });
    write_blob({
        path: "exceljs.xlsx",
        directory: Directory.Documents,
        blob: blob,
        fast_mode: true,
        recursive: true,
        on_fallback(error) {
          console.log("Xlsx no written.");
            console.error(error);
        }
    }).then(function () {
        console.log("xslx written.");
    });
  }).catch((err) => {
    console.log('Something really bad happened writing the XlSX file.', err);
  });

  /*workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], { type: BLOL_TYPE });
    Filesystem.writeFile({
      path: 'exceljs.xlsx',
      data: data.toString(),
      directory: Directory.Documents,
      encoding: Encoding.UTF8,
    }).then(response => {
      console.log("Successfully saved file")
    });
  }).catch((err) => {
    console.log('Something really bad happened writing the XlSX file.', err);
  });*/

  

  


  

  }

}
