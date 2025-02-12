// import ExcelJS from "exceljs";
// import { saveAs } from "file-saver";

// const ExamList = async () => {
//   const workbook = new ExcelJS.Workbook();
//   const worksheet = workbook.addWorksheet("Aptitude Training 2024");

//   // Merge and style main headings
//   worksheet.mergeCells("A1:J1");
//   worksheet.getCell("A1").value = "MBA";
//   worksheet.getCell("A1").alignment = { horizontal: "center" };
//   worksheet.getCell("A1").font = { bold: true, size: 14 };

//   worksheet.mergeCells("A2:J2");
//   worksheet.getCell("A2").value = "Aptitude Training 2024";
//   worksheet.getCell("A2").alignment = { horizontal: "center" };
//   worksheet.getCell("A2").font = { bold: true, size: 14 };

//   worksheet.mergeCells("A3:J3");
//   worksheet.getCell("A3").value = "Batch 2023-2025";
//   worksheet.getCell("A3").alignment = { horizontal: "center" };
//   worksheet.getCell("A3").font = { bold: true, size: 14 };

//   // Merge headers across rows 4 and 5
//   worksheet.mergeCells("A4:A5");
//   worksheet.getCell("A4").value = "Sl No";
//   worksheet.mergeCells("B4:B5");
//   worksheet.getCell("B4").value = "Registration Number";
//   worksheet.mergeCells("C4:C5");
//   worksheet.getCell("C4").value = "Student Name";
//   worksheet.mergeCells("D4:D5");
//   worksheet.getCell("D4").value = "Seat";

//   worksheet.getCell("E4").value = "10:30 AM\n11:40 AM";
//   worksheet.getCell("E4").alignment = { horizontal: "center", wrapText: true };
//   worksheet.getCell("E5").value = "Aug 13";
//   worksheet.getCell("E5").alignment = { horizontal: "center" };

//   worksheet.getCell("F4").value = "8:30 AM\n10:20 AM";
//   worksheet.getCell("F4").alignment = { horizontal: "center", wrapText: true };
//   worksheet.mergeCells("F5:G5");
//   worksheet.getCell("F5").value = "Aug 14";
//   worksheet.getCell("F5").alignment = { horizontal: "center" };

//   worksheet.getCell("G4").value = "10:30 AM\n11:40 AM";
//   worksheet.getCell("G4").alignment = { horizontal: "center", wrapText: true };

//   worksheet.mergeCells("H4:H5");
//   worksheet.getCell("H4").value = "Total Sessions";

//   worksheet.mergeCells("I4:I5");
//   worksheet.getCell("I4").value = "TA";

//   worksheet.mergeCells("J4:J5");
//   worksheet.getCell("J4").value = "Att %";

//   // Style for merged headers
//   const mergedHeaders = ["A4", "B4", "C4", "D4", "E4","E5", "F4","F5", "G4", "H4", "I4", "J4"];
//   mergedHeaders.forEach((cell) => {
//     worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
//     worksheet.getCell(cell).font = { bold: true };
//     worksheet.getCell(cell).fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: "FFFF00" } // Yellow background
//     };
//   });

//   // Sample Data
//   const data = [[1, "P19DU23M015001", "AISHWARYA SUSAN EBY", "MGMT", 1, 0, 0, 30, 22, 44]];

//   data.forEach((row, rowIndex) => {
//     const newRow = worksheet.addRow(row);

//     newRow.eachCell((cell, colIndex) => {
//       if (colIndex >= 5 && colIndex <= 7) {
//         // Apply conditional formatting
//         if (cell.value === 0) {
//           cell.fill = {
//             type: "pattern",
//             pattern: "solid",
//             fgColor: { argb: "FF0000" } // Red for 0
//           };
//           cell.font = { bold: true, color: { argb: "FFFFFF" } }; // White text
//         } else if (cell.value === 1) {
//           cell.fill = {
//             type: "pattern",
//             pattern: "solid",
//             fgColor: { argb: "00FF00" } // Green for 1
//           };
//           cell.font = { bold: true };
//         }
//       }
//     });
//   });

//   // Set Column Widths
//   worksheet.columns = [
//     { width: 6 }, { width: 20 }, { width: 25 }, { width: 10 },
//     { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 },
//     { width: 12 }, { width: 10 }, { width: 10 }
//   ];

//   // Add Borders
//   worksheet.eachRow({ includeEmpty: true }, (row) => {
//     row.eachCell((cell) => {
//       cell.border = {
//         top: { style: "thin" },
//         left: { style: "thin" },
//         bottom: { style: "thin" },
//         right: { style: "thin" }
//       };
//     });
//   });

//   // Save File
//   const buffer = await workbook.xlsx.writeBuffer();
//   const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
//   saveAs(blob, "Aptitude_Training_2024.xlsx");
// };

// export default ExamList;

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import fetchData from "../Utility/FechData";


const ExamList = async () => {
  const { headings, students } = fetchData(); // Fetch data from the mock utility
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Aptitude Training 2024");

  // Merge and style main headings from fetched data
  worksheet.mergeCells("A1:J1");
  worksheet.getCell("A1").value = headings.mainTitle;
  worksheet.getCell("A1").alignment = { horizontal: "center" };
  worksheet.getCell("A1").font = { bold: true, size: 14 };

  worksheet.mergeCells("A2:J2");
  worksheet.getCell("A2").value = headings.subTitle;
  worksheet.getCell("A2").alignment = { horizontal: "center" };
  worksheet.getCell("A2").font = { bold: true, size: 14 };

  worksheet.mergeCells("A3:J3");
  worksheet.getCell("A3").value = headings.batch;
  worksheet.getCell("A3").alignment = { horizontal: "center" };
  worksheet.getCell("A3").font = { bold: true, size: 14 };

  // Merge headers dynamically from the fetched data
  headings.columns.forEach((column, index) => {
    const cell = String.fromCharCode(65 + index) + "4"; // Convert index to column (A, B, C, ...)
    worksheet.getCell(cell).value = column;
  });

  // Style headers with yellow background and center alignment
  const headerCells = [
    "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4", "J4"
  ];
  
  headerCells.forEach((cell) => {
    worksheet.getCell(cell).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell(cell).font = { bold: true };
    worksheet.getCell(cell).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow background
    };
  });

  // Add student data dynamically from the fetched data
  students.forEach((student) => {
    const newRow = worksheet.addRow(student);

    // Apply conditional formatting for specific cells
    newRow.eachCell((cell, colIndex) => {
      if (colIndex >= 5 && colIndex <= 7) {
        if (cell.value === 0) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000" } // Red for 0
          };
          cell.font = { bold: true, color: { argb: "FFFFFF" } }; // White text
        } else if (cell.value === 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "00FF00" } // Green for 1
          };
          cell.font = { bold: true };
        }
      }
    });
  });

  // Set Column Widths
  worksheet.columns = [
    { width: 6 }, { width: 20 }, { width: 25 }, { width: 10 },
    { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 },
    { width: 12 }, { width: 10 }, { width: 10 }
  ];

  // Add Borders
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    });
  });

  // Save File
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, "Aptitude_Training_2024.xlsx");
};

export default ExamList;

