# Maaser Loader

This Node.js application reads data from an Excel file and loads it into MongoDB for tracking Maaser (tithing) data.

## Setup

1. Install dependencies:
```bash
npm install
```

2. Make sure MongoDB is running locally on port 27017

3. Place your Excel file (maaser-1.xlsx) in the root directory

4. Run the application:
```bash
npm start
```

## Data Structure

The application expects the Excel file to have:
- One tab per year
- Each tab contains monthly sections
- Each month has two charts:
  1. Income chart (source, earnings, date)
  2. Tzedaka chart (tzedaka, destination, date)

## MongoDB Collections

The data is stored in two collections:
1. `incomes` - Stores income data
2. `tzedakas` - Stores tzedaka (charity) data

## Note

Make sure to update the `TEMP_USER_ID` in `index.js` with your actual user ID before running the application.