const express = require("express");
const app = express();
const multer = require("multer");
const XLSX = require("xlsx");
const { MongoClient } = require("mongodb");
const path = require("path");

app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: "true" }));

const upload = multer({ dest: "uploads/" });

const uri = "mongodb://127.0.0.1:27017/excel_to_mongodb";
const client = new MongoClient(uri);

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.post("/upload", upload.single("file"), async (req, res) => {
  const filePath = req.file.path;

  // Read the Excel file
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);

  console.log("Parsed Excel data:", data);

  try {
    await client.connect();
    console.log("Connected successfully to MongoDB");
    const database = client.db("excel_to_mongodb"); // Replace with your database name
    const collection = database.collection("excel_data"); // Replace with your collection name

    // Insert data into MongoDB
    await collection.insertMany(data);
    console.log(`Inserted ${result.insertedCount} documents`);

    res.send("Data imported successfully!");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error importing data");
  } finally {
    await client.close();
  }
});

app.listen(3000, () => {
  console.log("Server is running on port 3000");
});
