# Business Analysis

This repository contains tools and formulas for business analysis tasks.

## Zingers Sales Tracking

Track monthly Zingers sales using Excel formulas.

### Quick Start

**Simple copy-paste formula for January 2024:**
```excel
=SUMIFS(C:C, B:B, "Zingers", A:A, ">=01/01/2024", A:A, "<=01/31/2024")
```

### Files

- **[FORMULAS.md](FORMULAS.md)** - Quick reference with ready-to-use formulas for each month
- **[zingers-sales-tracking.md](zingers-sales-tracking.md)** - Complete guide with methods, examples, and tips
- **[sample-data.csv](sample-data.csv)** - Sample data for testing formulas

### Data Structure

Your Excel sheet should have:
- **Column A**: Date (transaction date)
- **Column B**: Product Name
- **Column C**: Quantity Sold

### Available Methods

1. **SUMIFS** - Filter by product name and date range
2. **SUMPRODUCT** - Advanced multi-criteria calculation
3. **Pivot Tables** - Visual monthly breakdown
4. **Dynamic Formulas** - Automatically adjust to selected month

See the full documentation for detailed examples and troubleshooting.