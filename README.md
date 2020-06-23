

## Getting Started

### Install all dependencies
npm install

### Run the development environment:
npm run dev

### Run the Production environment:
npm run prod

Open [http://localhost:8080](http://localhost:8080) with your browser to see the result.

### Features:
1. Add/remove rows or columns to/from the existing sheet.
2. Editable and readonly cell.
3. Excel-like syntax (formulas start with "=").
4. Support for arbitrary expressions (=A1+B2*C3).
5. Automatic localStorage persistence.

### Validation used

1. Row and column input fields are mandatory, if there is no excel sheet present.
2. Row and column can be added explicitly, after creating the excel sheet.
3. Store all the data to localstorage.
4. Remove all data fom localstorage and clear the excel sheet.

### Live Demo
Open [https://pushp1992.github.io/Create-Excel-sheet-vanilajs/](https://pushp1992.github.io/Create-Excel-sheet-vanilajs/)
