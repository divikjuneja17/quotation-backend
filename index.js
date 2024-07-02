const express = require("express");
const cors = require("cors");
const puppeteer = require("puppeteer");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post("/api/quotes", async (req, res) => {
  const quoteData = req.body;

  const htmlContent = `
  <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Quote</title>
      <style>
         body { font-family: Arial, sans-serif; }
        .header { text-align: center; font-size: 24px; color: #337ab7; }
        .subheader { text-align: center; font-size: 18px; margin-top: 10px; color: #337ab7; }
        .details { margin: 20px 0; }
        .details div { margin: 5px 0; }
        .table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        .table, .table th, .table td { border: 1px solid black; }
        .table th, .table td { padding: 10px; text-align: left; }
        .footer { text-align: center; margin-top: 30px; }
        .total { text-align: right; margin-top: 10px; }

        .quote-form {
  margin: 20px 100px;
  font-family: Arial, sans-serif;
}

.header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.header h1 {
  font-size: 36px;
  color: rgb(37 99 235);
  font-weight: bold;
}

.header .date {
  font-size: 20px;
  color: #ffff;
  font-weight: bold;
  background-color: rgb(37 99 235);
  padding: 15px;
}

h2,
h3 {
  color: rgb(37 99 235);
  margin-top: 20px;
  font-weight: bold;
}

.form-section {
  margin: 20px 0px;
}

.form-group {
  display: flex;
  flex-direction: column;
  margin-bottom: 10px;
}

.form-group label {
  font-weight: bold;
  margin-bottom: 5px;
}

.form-group input,
.form-group textarea {
  padding: 8px;
  font-size: 16px;
  border: 1px solid #ccc;
  border-radius: 4px;
}

.table {
  display: table;
  width: 100%;
  border-collapse: collapse;
}

.table-header,
.table-row {
  display: table-row;
}

.table-cell {
  display: table-cell;
  padding: 10px;
  border: 1px solid #ccc;
}

.table-header .table-cell {
  font-weight: bold;
  background-color: #f2f2f2;
}

.save-quote {
  background-color: #28a745;
  color: #fff;
  padding: 10px 20px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.save-quote:hover {
  background-color: #218838;
}

.add-new {
  background-color: #2d88f7;
  color: #fff;
  padding: 10px 20px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.error {
  color: rgb(255, 0, 0);
}

.mb-10 {
  margin-bottom: 10px;
}

.from-to-section {
  display: flex; 
  justify-content: space-between;
  background-color: rgb(37 99 235); 
  padding: 20px 10px; 
  color: white;
}

.main-section {
  display: flex; 
  justify-content: space-between;
  margin: 10px 0px;
}
      </style>
    </head>
    <body>
      
      <div class="quote-form">
        <div class="header">
          <h1>BOLTCARGO</h1>
          <div class="date">${quoteData.date}</div>
        </div>
      
        <h3>Quote Estimate</h3>

        <div class="from-to-section">
          <div>
            <label>From: ${quoteData.from}</label>
          </div>
          <div>
            <label>To: ${quoteData.to}</label>
          </div>
        </div>
      
        <div class="main-section">
            <div>
              <label>Customer: ${quoteData.customer}</label>
            </div>
            <div>
              <label>Free Time: ${quoteData.freeTime}</label>
            </div>
        </div>

        <div class="main-section">
            <div>
              <label>Customer ID: ${quoteData.customerId}</label>
            </div>
            <div>
              <label>Incoterms: ${quoteData.incoterms}</label>
            </div>
        </div>

        <div class="main-section">
            <div>
              <label>Validity: ${quoteData.validity}</label>
            </div>
            <div>
              <label>Sailing: ${quoteData.sailing}</label>
            </div>
        </div>

        <div class="main-section">
            <div>
              <label>Transit Time: ${quoteData.transitTime}</label>
            </div>
            <div>
              <label>Commodity: ${quoteData.commodity}</label>
            </div>
        </div>

        <table class="table" style="margin-top: 30px;">
            <thead>
              <tr>
                <th>Description</th>
                <th>Quantity</th>
                <th>Price</th>
                <th>Total</th>
              </tr>
            </thead>
            <tbody>

              ${quoteData.items
                .map(
                  (item) => `
                <tr>
                  <td>${item.description}</td>
                  <td>${item.quantity}</td>
                  <td>${item.price}</td>
                  <td>${item.total}</td>
                </tr>
              `
                )
                .join("")}
            </tbody>
          </table>

          <div style="display: flex; justify-content: space-between">
            <div class="form-section terms">
                <label style="font-weight: bold">Terms and Conditions:</label>
                <ul>
                    <li>All rates quoted are valid for 15 days.</li>
                    <li>40% payment should be done in advance.</li>
                    <li>No returns will be accepted after 20 days.</li>
                    <li>The remaining amount should be paid within 20 days of delivery.</li>
                </ul>
              </div>

              <div class="form-section terms">
                <div>Total: $${quoteData.total}</div>
              </div>
        </div>
      </div>
    </body>
    </html>
  `;

  try {
    console.log("Launching browser...");
    const browser = await puppeteer.launch({
      headless: true,
      executablePath:
        process.env.NODE_ENV === "production"
          ? process.env.PUPPETEER_EXECUTABLE_PATH
          : puppeteer.executablePath(),
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--single-process",
        "--no-zygote",
        // Add more args if necessary
      ],
    });
    console.log("Browser launched.");

    console.log("Creating new page...");
    const page = await browser.newPage();
    console.log("Page created.");

    // Set request interception to disable loading of unnecessary resources
    await page.setRequestInterception(true);
    page.on("request", (req) => {
      if (
        ["image", "stylesheet", "font", "script"].includes(req.resourceType())
      ) {
        req.abort();
      } else {
        req.continue();
      }
    });

    await page.setContent(htmlContent, { waitUntil: "load", timeout: 0 }); // Disable timeout when setting content
    const pdfBuffer = await page.pdf({ format: "A4" });

    await browser.close();

    res.set({
      "Content-Type": "application/pdf",
      "Content-Disposition": "attachment; filename=quote.pdf",
      "Content-Length": pdfBuffer.length,
    });

    res.send(pdfBuffer);

    // Add new row to Excel file
    const headerData = [
      { header: "Sales Person", key: "salesPerson" },
      { header: "Customer", key: "customer" },
      { header: "Imp/Exp", key: "impExp" },
      { header: "LCL/FCL/Weight", key: "lclFclWeight" },
      { header: "POL", key: "pol" },
      { header: "POD", key: "pod" },
      { header: "Quantity", key: "quantity" },
      { header: "Price", key: "price" },
      { header: "Total", key: "total" },
      { header: "Free Time", key: "freeTime" },
      { header: "Validity", key: "validity" },
      { header: "Commodity", key: "commodity" },
      { header: "Incoterms", key: "incoterms" },
      { header: "Transit Time", key: "transitTime" },
      { header: "Sailing", key: "sailing" },
    ];

    const filePath = path.join(__dirname, "quotes.xlsx");

    console.log(`Checking file existence at path: ${filePath}`);
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (fs.existsSync(filePath)) {
      console.log("File exists. Reading file...");
      await workbook.xlsx.readFile(filePath);
      worksheet = workbook.getWorksheet(1);
      console.log("File read successfully.");
    } else {
      console.log("File does not exist. Creating new workbook...");
      worksheet = workbook.addWorksheet("Quotes");
      worksheet.columns = headerData;
      console.log("New workbook created.");
    }

    const newRow = worksheet.addRow({
      salesPerson: quoteData.salesPerson,
      customer: quoteData.customer,
      impExp: quoteData.impExp,
      lclFclWeight: quoteData.lclFclWeight,
      pol: quoteData.pol,
      pod: quoteData.pod,
      quantity: quoteData.quantity,
      price: quoteData.price,
      total: quoteData.total,
      freeTime: quoteData.freeTime,
      validity: quoteData.validity,
      commodity: quoteData.commodity,
      incoterms: quoteData.incoterms,
      transitTime: quoteData.transitTime,
      sailing: quoteData.sailing,
    });

    newRow.commit();
    await workbook.xlsx.writeFile(filePath);
    console.log("New row added to Excel file.");
  } catch (error) {
    console.error("Error processing the request:", error);
    res.status(500).send("An error occurred while processing the request.");
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
