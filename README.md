# excel-sales-summary  📈
A simple VBA macro for Microsoft Excel that automates the generation of a Summary Report from raw sales data.

It takes a dataset containing products, quantities, per-unit prices, and dates, and produces a clean, aggregated summary including:
- Total Quantity per product
- Total Price per product
- Grand Total for all products

----

## Folder Structure

```bash
excel-sales-summary/
│
├── SalesData.xlsm           # Sample Excel file with the SalesData sheet
├── README.md                # Project overview and explanation
├── VBA_Module.txt           # The macro code
└── Screenshots/
    └── SummaryReport.png    # Screenshot of the generated summary report
``` 

## Sample

### Sample Input

`SalesData` Sheet:
```bash
Product     Quantity   Price     Date
-----------------------------------------
Widget A	  10	   25	     2025-08-01
Widget B	  5	       50	     2025-08-02
Widget A	  3	       25	     2025-08-04
Widget C	  20	   15	     2025-08-02
...	          ...	   ...	     ...
...	          ...	   ...	     ...
...	          ...	   ...	     ...

 ---------------------------
|  Generate Summary Button  |  <---- vba macro runs upon pressing this(have to configure)
 ---------------------------
```

### Sample Output

`SummaryReport` Sheet:
```bash
Product	     Total Quantity	     Total Price
----------------------------------------------
Widget A	      13	            325
Widget B	      5	                250
Widget C	      20	            300
Grand Total	      38	            875
```

## How Macro Works
- The macro loops through all rows in the `SalesData` sheet.
- It tracks unique products and aggregates their quantities and total prices.
- Creates a new worksheet called `SummaryReport` (deletes previous report if it exists).
- Writes aggregated data to the new sheet and adds a Grand Total row.
- Formats headers and auto-fits columns for readability.
