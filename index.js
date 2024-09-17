import express from "express";
import xl from "excel4node";

const app = express();

// Serve the download page with a button
app.get("/Excel", function (req, res) {
  res.send(`
        <html>
          <body>
            <h1>Excel Download</h1>
            <button onclick="downloadExcel()">Download Excel</button>
            <script>
              function downloadExcel() {
                window.location.href = '/downloadExcel';
              }
            </script>
          </body>
        </html>
      `);
});

// Create the Excel file and send it as a response
const download = (req, res) => {
  // Create a new instance of a Workbook
  const wb = new xl.Workbook();

  // Add a new worksheet to the workbook
  const ws = wb.addWorksheet("Large Data Sheet");

  // Create header style
  const headerStyle = wb.createStyle({
    font: {
      bold: true,
      color: "#FFFFFF",
      size: 12,
    },
    fill: {
      type: "pattern",
      patternType: "solid",
      bgColor: "#4CAF50", // Background color
      fgColor: "#4CAF50", // Foreground color
    },
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
  });

  // Create a style for data rows
  const dataStyle = wb.createStyle({
    font: {
      size: 12,
    },
    alignment: {
      horizontal: "left",
    },
  });

  // Example dataset (you can replace this with actual data)
  const data = [
    ["ID", "Name", "Email", "Age", "Date of Joining"],
    [1, "John Doe", "john@example.com", 25, new Date()],
    [2, "Jane Smith", "jane@example.com", 28, new Date()],
    [3, "Bob Johnson", "bob@example.com", 32, new Date()],
    // Add more data here as needed
  ];

  // Merge cells for a title
  ws.cell(1, 1, 2, 5, true) // Merge A1 to E1
    .string("Employee Data Report")
    .style({
      font: { bold: true, size: 16 },
      alignment: { horizontal: "center" },
    });

  // Set column headers with headerStyle
  const headers = data[0];
  headers.forEach((header, i) => {
    ws.cell(2, i + 1)
      .string(header)
      .style(headerStyle);
  });

  // Fill data into cells with dataStyle
  for (let row = 1; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      const value = data[row][col];
      const cell = ws.cell(row + 2, col + 1);

      // Handle different types of data
      if (typeof value === "number") {
        cell.number(value).style(dataStyle);
      } else if (value instanceof Date) {
        cell.date(value).style(dataStyle);
      } else if (typeof value === "boolean") {
        cell.bool(value).style(dataStyle);
      } else {
        cell.string(String(value)).style(dataStyle);
      }
    }
  }

  // Adjust column widths automatically
  ws.column(1).setWidth(10); // ID column
  ws.column(2).setWidth(25); // Name column
  ws.column(3).setWidth(30); // Email column
  ws.column(4).setWidth(10); // Age column
  ws.column(5).setWidth(20); // Date column

  // Send the file as a response
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=Employee_Report.xlsx"
  );

  wb.write("Employee_Report.xlsx", res);
};

// Route to handle Excel download
app.get("/downloadExcel", download);

app.listen(9000, () => {
  console.log("Server is listening on port 9000");
});
