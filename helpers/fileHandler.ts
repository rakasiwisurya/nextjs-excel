import * as ExcelJS from "exceljs";

export type TJsonData = Array<{
  [key: string]: any;
}>;

export interface IExcelJSStyle {
  numFmt?: string;
  font?: Partial<ExcelJS.Font>;
  alignment?: Partial<ExcelJS.Alignment>;
  protection?: Partial<ExcelJS.Protection>;
  border?: Partial<ExcelJS.Borders>;
  fill?: ExcelJS.Fill;
}

export type TExcelJSStyleItem =
  | IExcelJSStyle["alignment"]
  | IExcelJSStyle["font"]
  | IExcelJSStyle["alignment"]
  | IExcelJSStyle["protection"]
  | IExcelJSStyle["border"]
  | IExcelJSStyle["fill"];

export interface IJsonToExcel {
  filename: string;
  data: {
    [sheetName: string]: {
      json: TJsonData;
      headerCells?: IExcelJSStyle;
      dataCells?: IExcelJSStyle;
      cells?: IExcelJSStyle;
    };
  };
}

export interface IExcelRow {
  [key: string]: any;
}

export const jsonToExcel = ({ filename, data }: IJsonToExcel) => {
  const workbook = new ExcelJS.Workbook();

  Object.keys(data).forEach((sheetName) => {
    const jsonData = data[sheetName].json;
    const headerCells: any | undefined = data[sheetName].headerCells;
    const dataCells: any | undefined = data[sheetName].dataCells;
    const cells: any | undefined = data[sheetName].cells;

    const worksheet = workbook.addWorksheet(sheetName);

    // Add header row
    const headers = Object.keys(jsonData[0]);
    worksheet.addRow(headers);

    // Add data rows
    jsonData.forEach((obj) => {
      const values = Object.values(obj);
      worksheet.addRow(values);
    });

    // Apply styling to header row
    // worksheet.getRow(1).font = { bold: true };
    // worksheet.getRow(1).eachCell((cell) => {
    //   cell.fill = {
    //     type: "pattern",
    //     pattern: "solid",
    //     fgColor: { argb: "FF0000FF" }, // Blue color
    //   };
    //   cell.alignment = {
    //     horizontal: "center",
    //   };
    // });

    // Apply styling to header row if any
    if (headerCells) {
      worksheet.getRow(1).eachCell((cell: any) => {
        Object.keys(headerCells).forEach((headerCell) => {
          cell[headerCell] = headerCells[headerCell];
        });
      });
    }

    // Apply styling to data rows
    // worksheet.eachRow((row, rowNumber) => {
    //   if (rowNumber === 1) {
    //     row.eachCell((cell) => {
    //       // set cell setting
    //     });
    //   }
    // });

    // Apply styling to data rows
    if (dataCells) {
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber !== 1) {
          row.eachCell((cell: any) => {
            Object.keys(dataCells).forEach((dataCell) => {
              cell[dataCell] = dataCells[dataCell];
            });
          });
        }
      });
    }

    // Apply styling to all rows
    // worksheet.eachRow((row) => {
    //   row.eachCell((cell) => {
    //     cell.font = { size: 12 }; // Set font size to 12
    //   });
    // });

    // Apply styling to all rows
    if (cells) {
      worksheet.eachRow((row) => {
        row.eachCell((cell: any) => {
          Object.keys(cells).forEach((cellItem) => {
            cell[cellItem] = cells[cellItem];
          });
        });
      });
    }

    // Auto-fit column widths
    headers.forEach((header, index) => {
      const column = worksheet.getColumn(index + 1); // Columns are 1-indexed
      let maxLength = header.length;
      worksheet.eachRow({ includeEmpty: true }, (row) => {
        const cell = row.getCell(index + 1);
        if (cell.value) {
          const cellLength = cell.value.toString().length;
          if (cellLength > maxLength) {
            maxLength = cellLength;
          }
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength + 2; // Set a minimum width of 10
    });
  });

  // Convert workbook to XLSX file format
  const workbookBlob = workbook.xlsx.writeBuffer().then((buffer) => {
    return new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  });

  // Create object URL for the Blob
  workbookBlob.then((blob) => {
    const url = window.URL.createObjectURL(blob);

    // Create anchor element to trigger download
    const link = document.createElement("a");
    link.href = url;
    link.download = `${filename}.xlsx`;
    link.click();

    // Clean up object URL
    window.URL.revokeObjectURL(url);
  });
};

export const excelToJson = (file: File | null): Promise<IExcelRow[]> => {
  return new Promise((resolve, reject) => {
    if (file) {
      const reader = new FileReader();

      reader.onload = async (e: ProgressEvent<FileReader>) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(data);

          const sheet = workbook.worksheets[0]; // Get the first worksheet
          const jsonData: IExcelRow[] = [];

          sheet.eachRow((row, rowNum) => {
            if (rowNum !== 1) {
              // Skip header row
              const rowData: IExcelRow = {};
              row.eachCell((cell, colNum) => {
                const headerCell = sheet.getRow(1).getCell(colNum);
                const headerKey = headerCell?.value?.toString() ?? ""; // Use empty string as default key
                const cellValue = cell.value;
                rowData[headerKey] = cellValue;
              });
              jsonData.push(rowData);
            }
          });

          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => {
        reject(reader.error);
      };

      reader.readAsArrayBuffer(file);
    } else {
      reject(new Error("No file provided"));
    }
  });
};

// import * as XLSX from "xlsx";

// export const jsonToExcel = (jsonData: Record<string, any>[]) => {
//   // const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(jsonData);
//   // // Define cell styles for header cells
//   // const headerCellStyle = {
//   //   font: { bold: true }, // Set font weight to bold
//   //   fill: { fgColor: { rgb: "FF0000FF" } }, // Set foreground color to blue
//   //   border: {
//   //     // Set border for header cells
//   //     top: { style: "thin" },
//   //     bottom: { style: "thin" },
//   //     left: { style: "thin" },
//   //     right: { style: "thin" },
//   //   },
//   // };
//   // // Apply header cell styles to the range of cells representing headers
//   // ws["!cols"] = [{ wch: 20 }, { wch: 20 }, { wch: 20 }]; // Example: setting column widths
//   // ws["!rows"] = [{ hpt: 20 }]; // Example: setting row heights
//   // ws["!merges"] = [{ s: { c: 0, r: 0 }, e: { c: 2, r: 0 } }]; // Example: merging header cells
//   // // Apply header cell styles to the range of cells representing headers
//   // for (let colIndex = 0; colIndex < jsonData.length; colIndex++) {
//   //   const headerCellRef: string = XLSX.utils.encode_cell({ r: 0, c: colIndex });
//   //   ws[headerCellRef].s = headerCellStyle;
//   // }
//   // // Define cell styles for data cells
//   // const dataCellStyle = {
//   //   font: { sz: 12 }, // Set font size to 12
//   // };
//   // // Apply data cell styles to all data cells
//   // const range: XLSX.Range = XLSX.utils.decode_range(ws["!ref"] as string);
//   // if (range && range.s && range.e) {
//   //   for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex++) {
//   //     for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
//   //       const dataCellRef: string = XLSX.utils.encode_cell({
//   //         r: rowIndex,
//   //         c: colIndex,
//   //       });
//   //       ws[dataCellRef].s = dataCellStyle;
//   //     }
//   //   }
//   // }
//   // // Create a new workbook and add the worksheet
//   // const wb: XLSX.WorkBook = XLSX.utils.book_new();
//   // XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
//   // // Write the workbook to a file
//   // XLSX.writeFile(wb, "output.xlsx");
// };

// export const excelToJson = (file: File | null): Promise<unknown[]> => {
//   return new Promise((resolve, reject) => {
//     if (file) {
//       const reader = new FileReader();

//       reader.onload = (e: ProgressEvent<FileReader>) => {
//         const data = new Uint8Array(e.target?.result as ArrayBuffer);
//         const workbook = XLSX.read(data, { type: "array" });

//         const sheetName = workbook.SheetNames[0];
//         const sheet = workbook.Sheets[sheetName];

//         const jsonData = XLSX.utils.sheet_to_json(sheet);
//         resolve(jsonData);
//       };

//       reader.onerror = () => {
//         reject(reader.error);
//       };

//       reader.readAsArrayBuffer(file);
//     } else {
//       reject(new Error("No file provided"));
//     }
//   });
// };
