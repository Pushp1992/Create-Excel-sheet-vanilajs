

## Getting Started

### Install all dependencies
npm install

### Run the development environment:
npm run dev

### Run the Production environment:
npm run prod

Open [http://localhost:8080](http://localhost:8080) with your browser to see the result.

### Validation used

1. All Input fields are mandatory by Default.
2. Automatically Detect the Card Type while entering the Card Number.
3. CVV is validated based on the Card Type.
4. For now, CVV must be of 4 digit for  "AMERICAN EXPRESS" card. And 3 digit for rest of the cards.
5. Card Number length is set to 19 digit default for any Card.
6. Card Validity is checked for YEAR and MONTH too.
7. Card Holder Name is Mandatory.
8. Card Number will aitomatically indented with space every after 4 digit of entry.
9. validity date will be indented by space every 2 digit. 
10. For validity date (First 2 digit are always Month and last 2 digit are always year), which is validated at client.
11. Card can be Saved, fetched and removed.
12. Auto-tab: cursor should jump to next field on successful completion of
one field.

## Learn More
