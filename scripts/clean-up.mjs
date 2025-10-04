import { parse } from "csv-parse/sync";
import { stringify } from "csv-stringify/sync";
import { promises as fs } from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { keywords, accounts, filePaths } from "./constants.mjs";

// Utility to validate if a date string is valid
function isValidDate(dateString) {
  if (!dateString) return false;
  const [day, month, year] = dateString.split("-");
  const date = new Date(`${year}-${month}-${day}`);
  return !isNaN(date.getTime());
}

// Utility to format date strings into "yyyy-mm-dd"
function formatDate(dateString) {
  const parts = dateString.split("-");
  let date;

  if (parts[1]?.length === 3) {
    // Format: "dd-MMM-yy" (e.g., 17-Mar-25)
    const [day, month, year] = parts;
    date = new Date(`${day} ${month} 20${year}`);
  } else if (parts.length === 3) {
    // Format: "dd-mm-yyyy" (e.g., 28-02-2025)
    const [day, month, year] = parts;
    date = new Date(`${year}-${month}-${day}`);
  } else if (dateString.includes(" ")) {
    // Format: "dd MMM yyyy" (e.g., 23 Mar 2025)
    date = new Date(dateString);
  }

  if (isNaN(date.getTime())) {
    throw new Error(`Invalid date format: ${dateString}`);
  }
  date = date.toLocaleDateString();
  date = date
    .split("/")
    .map((part) => part.padStart(2, "0"))
    .join("-"); // Convert to "mm-dd-yyyy" format
  const [month, day, year] = date.split("-");
  date = `${year}-${month}-${day}`; // Convert to "yyyy-mm-dd"
  return date; // Return "yyyy-mm-dd"
}

// Get all CSV/XLS files from a folder
async function getCsvFiles(folderPath) {
  const files = await fs.readdir(folderPath);
  return files.filter((file) => file.endsWith(".xls") || file.endsWith(".csv")); // tab delimited xls and comma delimited csv
}

// Parse CSV content into an array of objects
function parseCsvContent(content, headers) {
  const transactionStart = content.indexOf(headers.tableStartsAt);
  if (transactionStart === -1) {
    throw new Error(`Table start header not found: ${headers.tableStartsAt}`);
  }
  const transactionData = content.slice(transactionStart);
  return parse(transactionData, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
    skip_records_with_error: true,
    relax_column_count: true,
    delimiter: headers.delimiter,
  });
}
// find the type of transaction type: income, expense, transfer
function tagTransactionType(records) {
  // if the amount is positive, it is income, else expense, if category is transfer, it is transfer
  records.forEach((record) => {
    if (record.Category === "transfer") {
      record.Type = "transfer";
    } else if (Number(record.Amount) < 0) {
      record.Type = "income";
    } else {
      record.Type = "expense";
    }
  });

  return records;
}

function cleanupDescription(description) {
  // Split the description by '/', remove extra spaces, filter out pure numbers and empty strings, then join back with '-'
  return description
    .split("/")
    .map((part) => part.trim())
    .filter((part) => part !== "" && isNaN(part))
    .join("-");
}
function currencyStringToNumber(currencyString) {
  // Remove any non-numeric characters (except for decimal point)
  const cleanedString = currencyString.replace(/[^0-9.-]+/g, "");
  // Convert to number
  return parseFloat(cleanedString);
}
// Categorize a transaction based on keywords
function categorizeTransaction(description, isCredit) {
  for (const [category, keywordsArray] of Object.entries(keywords)) {
    if (
      keywordsArray.some((keyword) =>
        description.toLowerCase().includes(keyword)
      )
    ) {
      if (
        !["salary", "dividend", "interest", "transfer", "investment"].includes(
          category
        ) &&
        isCredit
      ) {
        return "refund";
      }
      return category;
    }
  }
  // if no match is found, check if the amount is positive and mark it as refund
  if (isCredit) {
    return "refund";
  }

  if (description.toLowerCase().includes("p2m")) {
    return "others - Merchant Payment";
  }
  if (description.toLowerCase().includes("p2a")) {
    return "others - Account Payment";
  }

  return "others"; // Default category if no match is found
}

// Process a single CSV file
async function processCsvFile(filePath, bankName, accountOwner, headers) {
  const content = await fs.readFile(filePath, "utf-8");
  const records = parseCsvContent(content, headers);

  const processedRecords = records
    .filter((item) => isValidDate(item[headers.date])) // Skip invalid dates
    .map((item) => ({
      Date: formatDate(item[headers.date]),
      Amount: item[headers.debit]
        ? currencyStringToNumber(item[headers.debit])
        : currencyStringToNumber(item[headers.credit]) * -1,
      Category: categorizeTransaction(
        cleanupDescription(item[headers.description]),
        currencyStringToNumber(item[headers.credit]) > 0 // isCredit boolean
      ),
      Bank: bankName,
      Owner: accountOwner,
      Description: cleanupDescription(item[headers.description]),
    }));

  return processedRecords;
}

// Main function to process all CSV files in a folder
async function processAllCsvFiles() {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);
  const folderPath = path.join(__dirname, filePaths.statements);
  const outputFolderPath = path.join(__dirname, filePaths.output);

  const csvFiles = await getCsvFiles(folderPath);
  let allRecords = [];

  for (const file of csvFiles) {
    const filePath = path.join(folderPath, file);
    const content = await fs.readFile(filePath, "utf-8");

    let headers, bankName, accountOwner;
    for (let account of accounts) {
      if (content.includes(account.accountNumber)) {
        headers = account.headers;
        bankName = account.bank;
        accountOwner = account.owner;
        break;
      }
    }
    // Skip files with unknown account types
    if (!headers) {
      console.log(`Unknown account type for file: ${file}`);
      continue;
    }

    const processedRecords = await processCsvFile(
      filePath,
      bankName,
      accountOwner,
      headers
    );
    allRecords = allRecords.concat(processedRecords);
  }
  const taggedRecords = tagTransactionType(allRecords);
  // Sort records by date
  const sortedRecords = taggedRecords.sort(
    (a, b) => new Date(a.Date) - new Date(b.Date)
  );

  // Write the combined records to a single output CSV file
  const outputCsvPath = path.join(outputFolderPath, filePaths.outputCsv);
  const csvContent = stringify(sortedRecords, { header: true });
  await fs.writeFile(outputCsvPath, csvContent, "utf-8");

  console.log(`Output written to ${outputCsvPath}`);

  // Write the combined records to a single output json file
  const outputJsonPath = path.join(outputFolderPath, filePaths.outputJson);
  const jsonContent = JSON.stringify(sortedRecords, null, 2);
  await fs.writeFile(outputJsonPath, jsonContent, "utf-8");
  console.log(`Output written to ${outputJsonPath}`);
}

// Run the main function
processAllCsvFiles().catch((error) => {
  console.error("Error processing files:", error.message);
});
