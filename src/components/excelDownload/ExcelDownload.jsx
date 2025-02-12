
import React from "react";
import CandidateListDownload from "../CandidateListDownload";
import ExamList from "../ExamList";
import  "./ExcelDownload.css";


const headingData = {
  collegeName: "Kurupanidhi College of Management",
  department: "Department of Training and Placements",
  date: "21/09/2024",
};


// Example data from an external source (API or static data)
const candidateData = [
  [
    1,
    "example1@gmail.com",
    "John",
    "Doe",
    "John Doe",
    "john.doe@example.com",
    "1234567890",
    "BBA",
    "Business Management",
    "2024",
    "No",
    "California",
    "ACCA",
    "IT Service",
    "No",
    "Yes",
  ],
  [
    2,
    "example2@gmail.com",
    "Jane",
    "Doe",
    "Jane Doe",
    "jane.doe@example.com",
    "0987654321",
    "MBA",
    "Business Management",
    "2023",
    "Yes",
    "New York",
    "CFA",
    "Finance",
    "Yes",
    "No",
  ],
];

const ExcelDownload = () => {
  return (
    <div className="excel-download-container">
    <div className="header">
      <h1>Candidate List Download</h1>
      <CandidateListDownload candidateData={candidateData} headingData={headingData} />
    </div>
  
    <div className="exam-list-section">
      <h1>Exam List Download</h1>
      <button className="download-button" onClick={ExamList}>Download Excel</button>
    </div>
  </div>
  
  );
};

export default ExcelDownload;

