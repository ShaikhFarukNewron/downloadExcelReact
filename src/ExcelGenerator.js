import React from "react";
import ExcelJS from "exceljs";

function ExcelGenerator() {
  const generateExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet1");

    // Define your data
    const data = [
      ["Name", "Age", "Email"],
      ["John Doe", 30, "john@example.com"],
      ["Jane Smith", 25, "jane@example.com"]
    ];

    // Add data to the worksheet
    data.forEach((row) => {
      worksheet.addRow(row);
    });

    // Apply some basic styling to the worksheet
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, size: 14, color: { argb: "FFFFFF" } }; // Bold, white text for header row
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "0070C0" }
    }; // Blue background color

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "center" }; // Center-align text vertically and horizontally
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        }; // Add thin borders to each cell
      });
    });

    // Generate the Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const filename = "data.xlsx";

    // Create a download link and trigger the download
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  return (
    <div className="p-4">
      <button
        className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded"
        onClick={generateExcel}
      >
        Generate Excel
      </button>
    </div>
  );
}

export default ExcelGenerator;
