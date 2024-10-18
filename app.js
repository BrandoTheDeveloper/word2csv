const express = require("express");
const mammoth = require("mammoth");
const multer = require("multer");  // Import multer for file uploads
const createCsvWriter = require("csv-writer").createObjectCsvWriter;
const path = require("path");
const fs = require("fs");

const app = express();
const port = 3000;

// Setup multer for handling file uploads
const upload = multer({ dest: "uploads/" });  // Files will be stored in 'uploads' folder

// Serve a simple upload form on the root route
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "index.html"));  // Serve the HTML file for upload form
});

// Handle the file upload and process the Word document
app.post("/upload", upload.single("docxFile"), express.urlencoded({ extended: true }), async (req, res) => {
    const docxFile = req.file.path;  // Path to the uploaded .docx file
    const csvFile = "output.csv";    // Path for the output CSV file

    try {
        // Call the function to parse Word document and generate CSV
        await parseDataToCsv(docxFile, csvFile);
        
        // Send the CSV file for download
        res.download(csvFile, (err) => {
            if (err) {
                console.error("Error sending the CSV file:", err);
                res.status(500).send("Failed to download the CSV file.");
            } else {
                // Clean up the uploaded Word document after processing
                fs.unlinkSync(docxFile);  // Optional: Delete uploaded .docx file after processing
                console.log("File processed and sent for download.");
            }
        });
    } catch (error) {
        console.error("Error processing the file:", error);
        res.status(500).send("Error processing the file.");
    }
});

// Function to extract text from the Word document
async function extractTextFromDocx(docxFile) {
    const result = await mammoth.extractRawText({ path: docxFile });
    return result.value;  // Returns the extracted text from the Word document
}

// Function to parse the extracted text and write to a CSV file
async function parseDataToCsv(docxFile, csvFile) {
    const text = await extractTextFromDocx(docxFile);  // Extract text from the Word document
    const lines = text.split("\n").filter(line => line.trim() !== "");  // Split lines by new line

    // Define the CSV writer with appropriate headers
    const csvWriter = createCsvWriter({
        path: csvFile,
        header: [
            { id: "potential_surplus", title: "Potential Surplus" },
            { id: "resale_value", title: "Est. Resale Value" },
            { id: "opening_bid", title: "Opening Bid" },
            { id: "date_sold", title: "Date Sold" },
            { id: "case_number", title: "Case #" },
            { id: "parcel_id", title: "Parcel ID" },
            { id: "type_of_foreclosure", title: "Type of Foreclosure" },
            { id: "first_name", title: "First Name" },
            { id: "last_name", title: "Last Name" },
            { id: "mailing_address", title: "Mailing Address" },
            { id: "mailing_city", title: "Mailing City" },
            { id: "mailing_state", title: "Mailing State" },
            { id: "mailing_zip_code", title: "Mailing Zip Code" },
            { id: "property_address", title: "Property Address" },
            { id: "property_city", title: "Property City" },
            { id: "property_state", title: "Property State" },
            { id: "property_zip_code", title: "Property Zip Code" },
            { id: "county", title: "County" }
        ]
    });

    // Map lines of text to CSV rows
    const records = lines.map(line => {
        const fields = line.split(",").map(field => field.trim());  // Split the line by commas
        return {
            potential_surplus: fields[0],
            resale_value: fields[1],
            opening_bid: fields[2],
            date_sold: fields[3],
            case_number: fields[4],
            parcel_id: fields[5],
            type_of_foreclosure: fields[6],
            first_name: fields[7],
            last_name: fields[8],
            mailing_address: fields[9],
            mailing_city: fields[10],
            mailing_state: fields[11],
            mailing_zip_code: fields[12],
            property_address: fields[13],
            property_city: fields[14],
            property_state: fields[15],
            property_zip_code: fields[16],
            county: fields[17]
        };
    });

    // Write the records to the CSV file
    await csvWriter.writeRecords(records);
    console.log("CSV file created successfully.");
}

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
