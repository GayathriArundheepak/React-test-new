// Mock utility function to simulate API response
const fetchData = () => {
    // Mocking data for the headings
    const headings = {
      mainTitle: "MBA",
      subTitle: "Aptitude Training 2024",
      batch: "Batch 2023-2025",
      columns: [
        "Sl No", "Registration Number", "Student Name", "Seat", "10:30 AM 11:40 AM", "8:30 AM 10:20 AM", "10:30 AM 11:40 AM", "Total Sessions", "TA", "Att %",
      ],
    };
  
    // Mocking data for the student entries
    const students = [
      [1, "P19DU23M015001", "AISHWARYA SUSAN EBY", "MGMT", 1, 0, 0, 30, 22, 44],
      [2, "P19DU23M015002", "JOHN DOE", "TECH", 0, 1, 1, 20, 18, 90],
      [3, "P19DU23M015003", "JANE SMITH", "HR", 1, 0, 1, 25, 20, 80],
    ];
  
    return { headings, students };
  };

  export default fetchData;
