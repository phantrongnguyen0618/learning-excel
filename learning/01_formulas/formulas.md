# Excel Formulas Complete Guide

## Introduction
In Microsoft Excel, formulas are expressions used to calculate, manipulate, and analyze data automatically.

All formulas begin with:

```excel
=
```

Example:

```excel
=2+3
```

Result:

```excel
5
```

---

# 1. Basic Formula Structure

```excel
= Value / Cell / Function / Operator
```

Example:

```excel
=A2*2
```

---

# 2. Cell References

| Formula | Meaning |
|---|---|
| `=A1` | Get value from cell A1 |
| `=A1+B1` | Add two cells |
| `=A1:A10` | Range from A1 to A10 |

---

# 3. Operators

| Operator | Meaning |
|---|---|
| `+` | Addition |
| `-` | Subtraction |
| `*` | Multiplication |
| `/` | Division |
| `^` | Power |
| `&` | Concatenate text |

Example:

```excel
=A1&B1
```

---

# 4. Basic Math Functions

## SUM

```excel
=SUM(A1:A10)
```

## AVERAGE

```excel
=AVERAGE(A1:A10)
```

## MAX

```excel
=MAX(A1:A10)
```

## MIN

```excel
=MIN(A1:A10)
```

## COUNT

```excel
=COUNT(A1:A10)
```

## COUNTA

```excel
=COUNTA(A1:A10)
```

---

# 5. Logical Functions

## IF

```excel
=IF(A1>=5,"Pass","Fail")
```

## AND

```excel
=AND(A1>5,B1>5)
```

## OR

```excel
=OR(A1>5,B1>5)
```

## NOT

```excel
=NOT(A1>5)
```

---

# 6. Lookup Functions

## XLOOKUP

```excel
=XLOOKUP(E2,A2:A10,B2:B10)
```

## VLOOKUP

```excel
=VLOOKUP(E2,A2:C10,2,FALSE)
```

## HLOOKUP

```excel
=HLOOKUP(E2,A1:J3,2,FALSE)
```

## INDEX

```excel
=INDEX(B2:B10,3)
```

## MATCH

```excel
=MATCH(E2,A2:A10,0)
```

## INDEX + MATCH

```excel
=INDEX(B2:B10,MATCH(E2,A2:A10,0))
```

---

# 7. Text Functions

## LEFT

```excel
=LEFT(A1,3)
```

## RIGHT

```excel
=RIGHT(A1,2)
```

## MID

```excel
=MID(A1,2,4)
```

## LEN

```excel
=LEN(A1)
```

## CONCAT

```excel
=CONCAT(A1,B1)
```

## TEXTJOIN

```excel
=TEXTJOIN(", ",TRUE,A1:A5)
```

## UPPER

```excel
=UPPER(A1)
```

## LOWER

```excel
=LOWER(A1)
```

## PROPER

```excel
=PROPER(A1)
```

## TRIM

```excel
=TRIM(A1)
```

## SUBSTITUTE

```excel
=SUBSTITUTE(A1,"old","new")
```

---

# 8. Date & Time Functions

## TODAY

```excel
=TODAY()
```

## NOW

```excel
=NOW()
```

## YEAR

```excel
=YEAR(A1)
```

## MONTH

```excel
=MONTH(A1)
```

## DAY

```excel
=DAY(A1)
```

## DATEDIF

```excel
=DATEDIF(A1,B1,"Y")
```

## NETWORKDAYS

```excel
=NETWORKDAYS(A1,B1)
```

---

# 9. Conditional Functions

## SUMIF

```excel
=SUMIF(A:A,"Laptop",B:B)
```

## SUMIFS

```excel
=SUMIFS(C:C,A:A,"Laptop",B:B,">100")
```

## COUNTIF

```excel
=COUNTIF(A:A,"IT")
```

## COUNTIFS

```excel
=COUNTIFS(A:A,"IT",B:B,">5")
```

## AVERAGEIF

```excel
=AVERAGEIF(A:A,"Laptop",B:B)
```

---

# 10. Error Handling Functions

## IFERROR

```excel
=IFERROR(A1/B1,0)
```

## ISERROR

```excel
=ISERROR(A1/B1)
```

## ISBLANK

```excel
=ISBLANK(A1)
```

---

# 11. Dynamic Array Functions

## FILTER

```excel
=FILTER(A1:C10,B1:B10="IT")
```

## UNIQUE

```excel
=UNIQUE(A1:A10)
```

## SORT

```excel
=SORT(A1:A10)
```

## SORTBY

```excel
=SORTBY(A1:B10,B1:B10,-1)
```

## SEQUENCE

```excel
=SEQUENCE(10)
```

---

# 12. Financial Functions

## PMT

```excel
=PMT(5%/12,60,-100000)
```

## FV

```excel
=FV(5%,10,-1000)
```

## PV

```excel
=PV(5%,10,-1000)
```

---

# 13. Statistical Functions

## MEDIAN

```excel
=MEDIAN(A1:A10)
```

## MODE

```excel
=MODE(A1:A10)
```

## STDEV

```excel
=STDEV(A1:A10)
```

## VAR

```excel
=VAR(A1:A10)
```

---

# 14. Reference Types

## Relative Reference

```excel
=A1+B1
```

## Absolute Reference

```excel
=$A$1
```

## Mixed Reference

```excel
=A$1
=$A1
```

---

# 15. Nested Formulas

Example:

```excel
=IF(AVERAGE(A1:A10)>=5,"Pass","Fail")
```

---

# 16. Common Excel Errors

| Error | Meaning |
|---|---|
| `#DIV/0!` | Division by zero |
| `#N/A` | Value not found |
| `#VALUE!` | Wrong data type |
| `#REF!` | Invalid reference |
| `#NAME?` | Wrong function name |

---

# 17. Most Important Formulas for Data Analysts

| Formula | Purpose |
|---|---|
| `SUMIFS` | Conditional summation |
| `COUNTIFS` | Conditional counting |
| `IF` | Logic |
| `IFERROR` | Handle errors |
| `XLOOKUP` | Advanced lookup |
| `INDEX MATCH` | Flexible lookup |
| `FILTER` | Dynamic filtering |
| `UNIQUE` | Remove duplicates |
| `SORT` | Dynamic sorting |
| `TEXT` | Format values |

---

# 18. Real-World Examples

## Total sales by product

```excel
=SUMIFS(C:C,A:A,"Laptop")
```

## Count IT employees

```excel
=COUNTIFS(B:B,"IT")
```

## Handle divide-by-zero

```excel
=IFERROR(A1/B1,0)
```

---

# 19. Recommended Learning Order

1. Basic arithmetic formulas
2. Cell references
3. SUM / AVERAGE
4. Logical functions
5. Text functions
6. Date functions
7. Lookup functions
8. Dynamic arrays
9. Data analysis formulas
10. Nested formulas

---

# 20. Formula vs Function

| Formula | Function |
|---|---|
| General expression | Built-in operation |
| `=A1+B1` | `=SUM(A1:B1)` |

---

# Conclusion

Excel formulas are the foundation of:
- Data Analysis
- Financial Modeling
- Business Reporting
- Dashboards
- Automation

Mastering formulas will significantly improve your productivity and analytical skills.