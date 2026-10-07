# PR. 1 Fundamental Booster: Excel Formulas Project

A hands-on Excel project covering the core formulas every data analyst uses day to day. It is built around three small business datasets (students, sales and employees) and shows each formula solving a real task: grading, discounts, lookups, filtering and date analysis.

📁 **File:** [`project_1.xlsx`](./project_1.xlsx)

> **Note:** XLOOKUP, XMATCH and FILTER need **Excel 365 / Excel 2021+** or **Google Sheets**.

---

## What's inside

| Sheet | Purpose |
|---|---|
| **Project Instructions** | Topics covered, main tasks and the deliverable |
| **Students Grade** | Grade classification, pass/fail logic, lookups and filtering |
| **Sales Data** | Pricing, discounts, GST, aggregations and a salesperson-by-month summary |
| **Employee Data** | Salary lookup, age and tenure, and text clean-up |

---

## Main tasks and how they were solved

### 1. Classify student grades (nested IF)
Students Grade, column F. Grade is based on the average of Math and Science.

```excel
=IF(E2>=90,"A",IF(E2>=75,"B",IF(E2>=60,"C",IF(E2>=40,"D","Fail"))))
```

### 2. Count students scoring above 60 (COUNTIFS)
```excel
=COUNTIFS(E2:E13,">60")
```

### 3. Find a product price (VLOOKUP)
Sales Data, column G. Pulls the price from the product table by product code.

```excel
=VLOOKUP(E2,$M$2:$O$6,3,FALSE)
```

### 4. Fetch employee salaries (XLOOKUP)
Employee Data, column F. Works even though the salary table is not sorted by ID.

```excel
=XLOOKUP(A2,$N$2:$N$7,$O$2:$O$7,"Not found")
```

### 5. Format and analyse the date of joining
Employee Data. Dates are formatted as `DD-MMM-YYYY`, then used to calculate age, tenure and days since joining.

```excel
=DATEDIF(D2,TODAY(),"Y")      // Age
=DATEDIF(E2,TODAY(),"Y")      // Tenure (years)
=TODAY()-E2                   // Days since joining
```

### 6. Extract top students (FILTER)
Students Grade, `L2:N13`. Returns every student whose average is above 80 and updates automatically when the data changes.

```excel
=FILTER(B2:F13,E2:E13>80)
```

---

## More formulas used

| Topic | Example | Where |
|---|---|---|
| IF + AND | `=IF(AND(C2>80,D2>80),"Yes","No")` | Students Grade, col G |
| IF + OR | `=IF(OR(H2>50000,F2>=10),"Eligible","Not eligible")` | Sales Data, col J |
| Nested IF (discount) | `=IF(H2>100000,10%,IF(H2>20000,5%,0))` | Sales Data, col I |
| Absolute vs relative reference | `=H2*$R$8` (GST rate locked in `$R$8`) | Sales Data, col K |
| SUMIFS | `=SUMIFS($H$2:$H$13,$B$2:$B$13,$Q2,$C$2:$C$13,R$1)` | Sales Data, `R2:T5` |
| AVERAGEIFS | `=AVERAGEIFS(E2:E13,E2:E13,">60")` | Students Grade |
| INDEX + MATCH | `=INDEX(R2:T5,MATCH("Riya",Q2:Q5,0),MATCH("Feb",R1:T1,0))` | Sales Data |
| XMATCH | `=XMATCH("P03",M2:M6)` | Sales Data |
| LEFT + FIND | `=LEFT(B2,FIND(" ",B2)-1)` | Employee Data, col J |
| UPPER / LOWER | `=UPPER(B2)`, `=LOWER(B2)` | Employee Data, cols K and L |
| ROUND / CEILING / FLOOR | `=ROUND(F2/12,2)` | Math functions |
| INDIRECT | `=SUM(INDIRECT("H2:H13"))` | Dynamic reference |
| OFFSET | `=SUM(OFFSET(R2,0,0,1,3))` | Dynamic range for sales trend |

---

## Sample results

- **Grades:** 4 students scored above 80 in both Math and Science.
- **Top performers (average > 80):** Aarav Shah, Isha Joshi, Yash Parmar, Hetvi Vyas, Pooja Rana.
- **Best salesperson:** Riya, with ₹5,10,000 in total sales across Jan to Mar.
- **Discount rules:** orders above ₹1,00,000 get 10%, above ₹20,000 get 5%.

---

## Skills demonstrated

- Relative and absolute cell references
- Conditional logic (IF, nested IF, AND, OR)
- Conditional aggregation (COUNTIFS, SUMIFS, AVERAGEIFS)
- Lookup functions (VLOOKUP, XLOOKUP, XMATCH, INDEX/MATCH)
- Text functions (LEFT, FIND, UPPER, LOWER)
- Date and time functions (DATEDIF, TODAY)
- Math functions (ROUND, CEILING, FLOOR)
- Dynamic references (INDIRECT, OFFSET)
- Dynamic arrays (FILTER)
- Data formatting (currency, dates, percentages)

---

## How to use

1. Download `project_1.xlsx`.
2. Open it in Excel 365 or upload it to Google Sheets.
3. Go through the sheets. Click any cell in a calculated column to see its formula.
4. Change any input (a score, a quantity, a price) and watch the dependent results update.

---

## Author

**[Your Name]**
[LinkedIn](https://linkedin.com/in/your-profile) · [GitHub](https://github.com/your-username)
