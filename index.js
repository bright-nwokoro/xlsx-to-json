import readXlsxFile from "read-excel-file/node";
import * as fs from "fs";
import * as path from "path";

const workBookLoc = "Book3.xlsx";

// Resolve the relative path to an absolute path
const absoluteWorkBookLoc = path.resolve(process.cwd(), workBookLoc);

// Readable Stream.
readXlsxFile(fs.createReadStream(absoluteWorkBookLoc)).then((rows) => {
    // Assuming the first row contains headers, convert the data to JSON
    const headers = rows[0];
    const jsonData = rows.slice(1).map(row => {
      const entry = {};
      headers.forEach((header, index) => {
        entry[header] = row[index];
      });
      return entry;
    });
  
    // Save the JSON data to a file or do something else with it
    const jsonFilePath = 'demo.json';
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2));
  
    console.log(`Conversion successful. JSON data saved to ${jsonFilePath}`);

    // console.log(rows[1]);
});
