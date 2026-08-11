# Lesson 03 — Sales Commission Analysis

## Project Overview

This project is part of my **Excel Data Analyst Journey**, where I am building practical Excel projects and documenting my progress publicly through GitHub and LinkedIn.

In this project, I built a hypothetical **Sales Commission & Target Performance Analysis** model to demonstrate how **relative, absolute, and mixed cell references** can be used to solve a practical business problem.

The model evaluates salesperson performance against sales targets and calculates performance-adjusted commissions.

---

## Business Problem

A hypothetical company wants to evaluate the performance of its salespeople and calculate their commissions based on their sales performance.

The analysis answers questions such as:

1. How much of each salesperson's sales target was achieved?
2. Did the salesperson perform below, at, or above their target?
3. What performance multiplier should be applied?
4. How much commission should each salesperson receive?
5. How can Excel formulas be designed so that they can be copied without manually changing cell references?

---

## Business Assumptions

Because this is a learning project and no actual company's commission policy was provided, the commission rates and performance multipliers used in this analysis are **hypothetical assumptions created specifically for this project**.

They do not represent the compensation policy of any real company.

### Commission Rate

A standard commission rate of **5%** was used for the final salesperson commission calculation.

### Performance Multipliers

| Performance Status | Multiplier | Explanation                          |
| ------------------ | ---------: | ------------------------------------ |
| Below Target       |        0.8 | 80% of the selected commission rate  |
| Target Achieved    |        1.0 | 100% of the selected commission rate |
| Above Target       |        1.2 | 120% of the selected commission rate |

### Performance Rules

|       Achievement | Performance Status |
| ----------------: | ------------------ |
|    Less than 100% | Below Target       |
|      Exactly 100% | Target Achieved    |
| Greater than 100% | Above Target       |

These assumptions were created to provide a realistic business scenario for demonstrating Excel cell references.

---

# Dataset

The dataset contains salesperson information including:

- Salesperson
- Region
- Sales
- Target

Additional calculated fields were created during the analysis:

- Target Achievement
- Performance Status
- Performance Multiplier
- Commission

Example:

| Salesperson |    Sales |   Target |
| ----------- | -------: | -------: |
| Chidinma    | ₦800,000 | ₦700,000 |
| James       | ₦910,000 | ₦800,000 |
| Yusuf       | ₦770,000 | ₦750,000 |
| Chika       | ₦540,000 | ₦750,000 |

---

# Key Calculations

## 1. Target Achievement

Target achievement was calculated by dividing Sales by Target.

Example excel formula:

=C2/D2

for Chidinma:

₦800,000 ÷ ₦7000,000 = 114.29%

for James:

₦9100,000 ÷ ₦800,000 = 113.75%

And the result was formatted as a percentage.

- Performance Status

An IF formula was used to classify each salesperson based on target achievement.

=IF(E2<100%,"Below Target",IF(E2=100%,"Target Achieved","Above Target"))

This was expected to produces three performance categories:

- Below Target
- Target Achieved
- Above Target

- Performance Multiplier

A performance multiplier was assigned based on the salesperson's performance status using the excel formula:

=IF(F2="Below Target",0.8,IF(F2="Target Achieved",1,1.2)).

This was expected to produce:
Below Target -> 0.8
Target Achieved -> 1.0
Above Target -> 1.2

4. Commission Calculation

The final commission formula used was:

Commission = Sales × Commission Rate × Performance Multiplier

The commission rate of 5% was stored in a separate cell so that it could be reused across the calculations.

=C2*$k$2*G2

Where:

C2 = Sales
$k$2 = Fixed 5% commission rate
G2 = Performance Multiplier

Example

For Chidinma

Sales = ₦800,000
Commission Rate = 5%
Performance Multiplier = 0.8

₦800,000 × 5% × 1.2
= ₦48,000

- Cell References Demonstrated

One of the main objectives of this project was to understand how different types of Excel cell references behave when formulas are copied.

Relative Cell References

Example: excel formula:
=C2/D2

Both references are relative. When the formula is copied down, Excel automatically changes:

=C2/D2
=C3/D3
=C4/D4
=C5/D5

This is useful when each row contains its own Sales and Target values.

- Absolute Cell References

Example:

=C2*$k$2*G2

The $k$2 reference is absolute.The dollar signs lock both:

Column k and Row 2

Therefore, when the formula is copied down, $k$2 remains unchanged.

This is useful because the 5% commission rate is a fixed busines assumption that should be applied consistently

- Mixed Cell References

A mixed-reference matrix was created to demonstrate how a row or column can be partially locked.

Example:

| Performance     | Multiplier |   3% |   5% |   7% |
| --------------- | ---------: | ---: | ---: | ---: |
| Below Target    |        0.8 | 2.4% | 4.0% | 5.6% |
| Target Achieved |        1.0 | 3.0% | 5.0% | 7.0% |
| Above Target    |        1.2 | 3.6% | 6.0% | 8.4% |

The formula used was: =$B2\*C$1
$B2

The column B is locked, but the row can change
$B2
$B3
$B4

C$1

The row 1 is locked, but the column can change.

C$1
D$1
E$1

This allows one formula to be copied across and down while Excel automatically adjusts the appropriate references.

- Key Excel Skills Demonstrated

This project demonstrates the following Excel skills:

Relative cell references
Absolute cell references
Mixed cell references
IF functions
Percentage calculations
Excel Tables
Structured references
Data organization
Business calculations
Formula copying
Performance classification

- Project Files

Sales_Commission_Analysis.xlsx Complete Excel analysis workbook
01_Raw_Sales_Data.png Raw dataset
02_Relative_Reference.png Relative reference demonstration
03_Absolute_Reference.png Absolute reference demonstration
04_Mixed_Reference.png Mixed reference demonstration
05_Business_Rule.png Commission calculation
06_Final_Analysis.png Final analysis and results

- Lessons Learned

The main lesson from this project is that Excel cell references control how formulas behave when they are copied.

Relative Reference: Changes automatically when copied.

Absolute Reference: Remains fixed when copied.

Mixed Reference: Locks either the row or the column while allowing the other part to change.

- Data Analyst Journey

This project is part of my ongoing Excel Data Analyst Journey, where I am building practical projects and documenting my progress publicly.

My learning roadmap is:

Excel -> SQL -> Python/Pandas -> Power BI

Learn -> Build -> Analyze -> Document -> Share -> Improve.

Author

Ifeoluwa Odunola
