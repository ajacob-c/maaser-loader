const ExcelJS = require('exceljs');
const mongoose = require('mongoose');
const Income = require('./models/Income');
const Tzedaka = require('./models/Tzedaka');
require('dotenv').config();

// Get user ID from environment variable
const USER_ID = process.env.USER_ID;

if (!process.env.MONGODB_URI) {
  console.error('MONGODB_URI is not defined in environment variables');
  process.exit(1);
}

if (!USER_ID) {
  console.error('USER_ID is not defined in environment variables');
  process.exit(1);
}

function parseDate(dateStr, monthYear) {
  if (!dateStr) return null;
  
  try {
    // If dateStr is just a month (e.g., "Nov"), use the 1st of the month
    if (dateStr.length <= 3) {
      const [month, year] = monthYear.split("'");
      const fullYear = '20' + year.trim();
      return new Date(`${fullYear}-${dateStr}-01`);
    }
    
    // If dateStr is in format "DD-MMM" (e.g., "16-Nov")
    const [day, monthAbbr] = dateStr.split('-');
    if (day && monthAbbr) {
      const [month, year] = monthYear.split("'");
      const fullYear = '20' + year.trim();
      return new Date(`${fullYear}-${monthAbbr}-${day}`);
    }
    
    // If dateStr is a full date string
    const date = new Date(dateStr);
    if (!isNaN(date.getTime())) {
      return date;
    }
    
    console.log('Failed to parse date:', { dateStr, monthYear });
    return null;
  } catch (error) {
    console.log('Error parsing date:', { dateStr, monthYear, error: error.message });
    return null;
  }
}

function parseAmount(amountValue) {
  if (!amountValue) return null;
  
  // If it's already a number, return it
  if (typeof amountValue === 'number') {
    return amountValue;
  }
  
  // If it's a string, clean it and parse it
  if (typeof amountValue === 'string') {
    // Remove currency symbols, commas, and any extra spaces
    const cleanAmount = amountValue.replace(/[$,]/g, '').trim();
    const parsedAmount = parseFloat(cleanAmount);
    return isNaN(parsedAmount) ? null : parsedAmount;
  }
  
  // If it's a cell object with a value property
  if (amountValue && typeof amountValue === 'object' && 'value' in amountValue) {
    return parseAmount(amountValue.value);
  }
  
  return null;
}

async function processExcelFile() {
  try {
    // Connect to MongoDB first
    await mongoose.connect(process.env.MONGODB_URI);
    console.log('Connected to MongoDB');
    
    // Read the Excel file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('maaser-1.xlsx');
    
    const incomePromises = [];
    const tzedakaPromises = [];
    
    // Process each sheet (year)
    for (const worksheet of workbook.worksheets) {
      let currentMonth = null;
      
      // Process each row
      worksheet.eachRow((row, rowNumber) => {
        // Skip empty rows
        if (!row || row.cellCount === 0) return;
        
        const rowData = row.values;
        
        // Check if this is a month header (e.g., "January '24")
        if (typeof rowData[1] === 'string' && rowData[1].includes("'")) {
          currentMonth = rowData[1].trim();
          return;
        }
        
        // Skip header rows and total rows
        if (typeof rowData[1] === 'string' && 
            (rowData[1].toLowerCase().includes('earnings') || 
             rowData[1].toLowerCase().includes('total') ||
             rowData[1].toLowerCase().includes('maaser') ||
             rowData[1].toLowerCase().includes('prev month'))) {
          return;
        }
        
        // Process income data if we have a valid row
        if (rowData[2] && rowData[3] && rowData[4]) {
          const amount = parseAmount(rowData[2]);
          const source = rowData[3].toString().trim();
          const date = rowData[4];
          
          if (amount && source && date instanceof Date) {
            const incomeData = {
              user: USER_ID,
              source,
              amount,
              date
            };
            incomePromises.push(
              Income.create(incomeData)
                .catch(err => console.error('Error creating income:', err))
            );
          }
        }
        
        // Process tzedaka data if we have a valid row
        if (rowData[7] && rowData[8] && rowData[9]) {
          const amount = parseAmount(rowData[7]);
          const organization = rowData[8].toString().trim();
          const date = rowData[9];
          
          if (amount && organization && date instanceof Date) {
            const tzedakaData = {
              user: USER_ID,
              organization,
              amount,
              date
            };
            tzedakaPromises.push(
              Tzedaka.create(tzedakaData)
                .catch(err => console.error('Error creating tzedaka:', err))
            );
          }
        }
      });
    }
    
    // Wait for all database operations to complete
    await Promise.all([...incomePromises, ...tzedakaPromises]);
    console.log('Data import completed successfully');
    
  } catch (error) {
    console.error('Error processing Excel file:', error);
  } finally {
    // Close MongoDB connection after all operations are complete
    await mongoose.connection.close();
    console.log('MongoDB connection closed');
  }
}

// Run the import
processExcelFile().catch(console.error); 