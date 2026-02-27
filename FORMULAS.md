# Quick Reference: Excel Formulas for Zingers Monthly Sales

## Basic Formula (Copy-Paste Ready)

### For a specific month in 2024:

**January 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=01/01/2024", A:A, "<=01/31/2024")
```

**February 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=02/01/2024", A:A, "<=02/29/2024")
```

**March 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=03/01/2024", A:A, "<=03/31/2024")
```

**April 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=04/01/2024", A:A, "<=04/30/2024")
```

**May 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=05/01/2024", A:A, "<=05/31/2024")
```

**June 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=06/01/2024", A:A, "<=06/30/2024")
```

**July 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=07/01/2024", A:A, "<=07/31/2024")
```

**August 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=08/01/2024", A:A, "<=08/31/2024")
```

**September 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=09/01/2024", A:A, "<=09/30/2024")
```

**October 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=10/01/2024", A:A, "<=10/31/2024")
```

**November 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=11/01/2024", A:A, "<=11/30/2024")
```

**December 2024:**
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=12/01/2024", A:A, "<=12/31/2024")
```

## Universal Formula Template

Replace `MM`, `DD`, and `YYYY` with your values:
```
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=MM/01/YYYY", A:A, "<=MM/DD/YYYY")
```

Where:
- First `MM` = Month number (01-12)
- `DD` = Last day of month (28, 29, 30, or 31)
- `YYYY` = Year

## Dynamic Formula (Automatic Month Detection)

If cell E2 contains any date in the month you want to analyze:
```
=SUMIFS($C:$C, $B:$B, "Zingers", $A:$A, ">="&DATE(YEAR(E2),MONTH(E2),1), $A:$A, "<"&DATE(YEAR(E2),MONTH(E2)+1,1))
```

## What These Formulas Assume

Your Excel data is organized as:
- **Column A**: Transaction dates
- **Column B**: Product names (including "Zingers")
- **Column C**: Quantity sold

## How to Use

1. Open your Excel file
2. Click on an empty cell where you want the result
3. Copy one of the formulas above
4. Paste it into the cell
5. Press Enter
6. The cell will show the total Zingers sold for that month

## Need Help?

See the full guide: [zingers-sales-tracking.md](zingers-sales-tracking.md)
