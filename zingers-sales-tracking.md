# Zingers Sales Tracking - Excel Formulas Guide

This guide provides Excel formulas to track and analyze Zingers sales per month.

## Data Structure

Your Excel sheet should have the following columns:
- **Column A**: Date (transaction date)
- **Column B**: Product Name
- **Column C**: Quantity Sold
- **Column D**: Month (optional, can be extracted from date)

Example:
```
| Date       | Product Name | Quantity Sold |
|------------|--------------|---------------|
| 01/05/2024 | Zingers      | 50            |
| 01/15/2024 | Zingers      | 30            |
| 02/03/2024 | Zingers      | 45            |
| 02/20/2024 | Other Product| 20            |
```

## Excel Formulas for Monthly Zingers Sales

### Method 1: SUMIFS Formula (Recommended)
Calculate total Zingers sold in a specific month:

```excel
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=01/01/2024", A:A, "<=01/31/2024")
```

**Explanation:**
- `C:C` - Sum values from Quantity Sold column
- `B:B, "Zingers"` - Where Product Name equals "Zingers"
- `A:A, ">=01/01/2024"` - Where Date is greater than or equal to month start
- `A:A, "<=01/31/2024"` - Where Date is less than or equal to month end

### Method 2: SUMIF with MONTH Function
If you want to sum all Zingers for a specific month across all years:

First, create a helper column (Column D) with this formula:
```excel
=MONTH(A2)
```

Then use:
```excel
=SUMIFS(C:C, B:B, "Zingers", D:D, 1)
```
Replace `1` with the month number (1=January, 2=February, etc.)

### Method 3: SUMPRODUCT (Advanced)
Calculate Zingers sold in January 2024:
```excel
=SUMPRODUCT((B2:B1000="Zingers")*(MONTH(A2:A1000)=1)*(YEAR(A2:A1000)=2024)*(C2:C1000))
```

**Explanation:**
- Checks if Product Name is "Zingers"
- Checks if Month is January (1)
- Checks if Year is 2024
- Multiplies by Quantity Sold

### Method 4: Pivot Table Summary
For a complete monthly breakdown:

1. Select your data range
2. Insert > PivotTable
3. Drag "Date" to Rows (group by Month)
4. Drag "Product Name" to Filters
5. Drag "Quantity Sold" to Values (Sum)
6. Filter for "Zingers"

## Monthly Summary Table

Create a summary table with formulas:

| Month      | Formula                                                                    | Result |
|------------|----------------------------------------------------------------------------|--------|
| January    | `=SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">=01/01/2024", $A:$A, "<=01/31/2024")` | Auto  |
| February   | `=SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">=02/01/2024", $A:$A, "<=02/29/2024")` | Auto  |
| March      | `=SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">=03/01/2024", $A:$A, "<=03/31/2024")` | Auto  |

### Dynamic Monthly Formula
For a more dynamic approach, if cell F2 contains the month/year (e.g., "01/2024"):

```excel
=SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">="&DATE(YEAR(F2),MONTH(F2),1), $A:$A, "<"&DATE(YEAR(F2),MONTH(F2)+1,1))
```

## Tips and Best Practices

1. **Use Named Ranges**: Define names for your columns (e.g., "Dates", "Products", "Quantities") for clearer formulas
2. **Absolute References**: Use `$` signs to lock cell references when copying formulas
3. **Data Validation**: Create a dropdown list of product names to ensure consistency
4. **Date Formatting**: Ensure dates are properly formatted as dates, not text
5. **Regular Updates**: Keep your data updated regularly for accurate monthly reporting

## Example Implementation

For a quick setup, create a summary section in your sheet:

```
Cell E1: "Month"
Cell F1: "Zingers Sold"

Cell E2: "January 2024"
Cell F2: =SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">=01/01/2024", $A:$A, "<=01/31/2024")

Cell E3: "February 2024"
Cell F3: =SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">=02/01/2024", $A:$A, "<=02/29/2024")
```

Continue for all months you want to track.

## Troubleshooting

- **#VALUE! error**: Check that your date column contains actual dates, not text
- **Wrong totals**: Verify that "Zingers" is spelled exactly the same in all rows
- **Missing data**: Ensure date ranges cover all days in the month
- **Case sensitivity**: Use `=SUMIFS($C:$C, $B:$B, "zingers")` for case-insensitive matching

## Additional Resources

For more complex analysis, consider:
- Creating charts to visualize monthly trends
- Using conditional formatting to highlight high/low sales months
- Adding year-over-year comparison formulas
