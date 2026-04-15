const XLSX = require("xlsx");
const fs = require("fs");

const file = XLSX.readFile("vvpat.xlsx");
const sheet = file.Sheets[file.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

fs.writeFileSync("vvpat.json", JSON.stringify(data, null, 2));

console.log("✅ JSON created successfully");