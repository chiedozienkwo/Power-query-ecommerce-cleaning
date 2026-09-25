# E-Commerce Order Data Cleaning — Power Query

Cleaned a 1,200-row e-commerce order dataset in Excel using Power Query, transforming raw, inconsistent data into a validated, analysis-ready table with an automated data-quality check built in.

## Overview

- **Tool:** Microsoft Excel — Power Query (M language)
- **Dataset:** 1,200 e-commerce orders, 14 columns (Order ID, Date, Customer ID, Product, Quantity, Unit Price, Shipping Address, Payment Method, Order Status, Tracking Number, Items in Cart, Coupon Code, Referral Source, Total Price)
- **Goal:** Standardize data types, resolve missing values, remove duplicates, and validate pricing logic — turning raw exported data into a trustworthy source for analysis or reporting

## Process

1. **Loaded the data** — `Data → From Table/Range` into the Power Query Editor
2. **Corrected data types** across all 14 columns (IDs and text fields as Text, `Quantity`/`ItemsInCart` as Whole Number, `UnitPrice`/`TotalPrice` as Decimal Number, `Date` as Date)
3. **Trimmed whitespace** from every text column in one operation (`Transform → Format → Trim`)
4. **Replaced blank values** in `CouponCode` with `"No Coupon"` (309 of 1,200 orders had no coupon applied)
5. **Removed duplicate rows** on `OrderID`
6. **Added a validation column**, `PriceCheck`, confirming `Quantity × UnitPrice = TotalPrice` for every row
7. **Loaded the result** back into Excel as a live, refreshable query — the whole pipeline re-runs automatically if the source data changes

### Power Query (M) — validation step

```
AddedValidation = Table.AddColumn(RemovedDuplicates, "PriceCheck", each
    if Number.Round([Quantity]*[UnitPrice],2) = [TotalPrice]
    then "OK" else "Mismatch")
```

## Debugging: an error along the way

After adding the `PriceCheck` column, every single row returned `Error` instead of `OK`/`Mismatch`. Power Query's error detail pane flagged the cause precisely:

> `Expression.Error: The field 'Quanty' of the record wasn't found.`

**Root cause:** a typo in the custom column formula — `[Quanty]` instead of `[Quantity]`.

**Fix:** reopened the `Added Custom` step via the Applied Steps pane, corrected the field reference, and re-ran the query. All 1,200 rows recalculated correctly, with 100% passing validation (no price mismatches).

This is a good illustration of why Power Query's step-based, case-sensitive M formulas are worth understanding directly rather than only through the UI — the Applied Steps panel and error detail messages make locating and fixing this kind of issue fast.

## Result

| Check | Result |
|---|---|
| Rows processed | 1,200 |
| Duplicate Order IDs removed | 0 (verified none existed) |
| Blank coupon codes resolved | 309 → labeled "No Coupon" |
| Price validation (`Quantity × UnitPrice = TotalPrice`) | 1,200 / 1,200 passed |

## Skills demonstrated

- Power Query (M) data transformation and type management
- Data quality validation logic (custom calculated columns)
- Debugging M-code errors using Applied Steps and error detail messages
- Building refreshable, repeatable ETL pipelines in Excel
  
- ## Before & After

**Raw data**


![Raw dataset before cleaning](Screenshot%202026-09-25%20123839_1.png)



**Cleaned data**


![Cleaned dataset with validation column](Screenshot%202026-09-25%20124132_1.png)

## Author

**Nkwo Chiedozie Jonathan**
[LinkedIn](https://linkedin.com/in/chiedozie-nkwo-b01912383) · [GitHub](https://github.com/chiedozienkwo)
