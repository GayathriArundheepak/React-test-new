
import React from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const CandidateListDownload = ({ candidateData, headingData }) => {
  const generateExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Candidate List");

    // Main Heading - Dynamic
    worksheet.mergeCells("A1:P1");
    worksheet.getCell("A1").value = headingData.collegeName;
    worksheet.getCell("A1").alignment = { horizontal: "center" };
    worksheet.getCell("A1").font = { bold: true, size: 14 };
    worksheet.getCell("A1").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow color
    };

    // Department Heading - Dynamic
    worksheet.mergeCells("A2:P2");
    worksheet.getCell("A2").value = headingData.department;
    worksheet.getCell("A2").alignment = { horizontal: "center" };
    worksheet.getCell("A2").font = { bold: true, size: 12 };
    worksheet.getCell("A2").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow color
    };

    // Date Heading - Dynamic
    worksheet.mergeCells("A3:P3");
    worksheet.getCell("A3").value = `Candidate List Date: ${headingData.date}`;
    worksheet.getCell("A3").alignment = { horizontal: "center" };
    worksheet.getCell("A3").font = { bold: true, size: 12 };
    worksheet.getCell("A3").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }, // Yellow color
    };



    // Table Headers
    const headers = [
      "Sl.No",
      "Email Address",
      "First Name",
      "Last Name",
      "Full Name",
      "Personal Email ID",
      "Phone Number",
      "PG Degree",
      "PG Degree Specialization",
      "PG Pass Out Year",
      "Active Backlog (Yes/No)",
      "Student Native Place (State)",
      "Professional Certification",
      "Service Line Applying For",
      "Do You Have Any Active Backlog?",
      "Sign",
    ];

    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: "center" };

    // Add Yellow Background to Headers
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF00" }, // Yellow color
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Append Data Rows
    candidateData.forEach((row) => {
      const rowData = worksheet.addRow(row);
      rowData.alignment = { horizontal: "center" };

      rowData.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Column Widths
    worksheet.columns = [
      { width: 6 }, // Sl.No
      { width: 25 }, // Email
      { width: 12 }, // First Name
      { width: 12 }, // Last Name
      { width: 20 }, // Full Name
      { width: 25 }, // Personal Email ID
      { width: 15 }, // Phone Number
      { width: 10 }, // PG Degree
      { width: 25 }, // PG Degree Specialization
      { width: 15 }, // PG Pass Out Year
      { width: 12 }, // Active Backlog
      { width: 15 }, // Student Native Place
      { width: 20 }, // Professional Certification
      { width: 20 }, // Service Line
      { width: 20 }, // Active Backlog Check
      { width: 10 }, // Sign
    ];

    // Generate Excel File
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "candidate_list.xlsx");
  };

  return (
    <div>
      <button  className="download-button" onClick={generateExcel}>Download Excel</button>
    </div>
  );
};

export default CandidateListDownload;
