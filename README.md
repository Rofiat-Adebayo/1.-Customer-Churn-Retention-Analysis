# YDP Bank | Customer Churn & Retention Analysis
An end-to-end BI project analysing customer churn at YDP Bank to identify at-risk customer segments, uncover churn patterns, and recommend data-driven retention strategies using Excel, Power Query, Power Pivot, DAX, and VBA.



<img width="1321" height="748" alt="image" src="https://github.com/user-attachments/assets/695f9f85-42d5-4485-87b1-c753ce42c773" />

---


<img width="1380" height="723" alt="image" src="https://github.com/user-attachments/assets/a6584ce7-ec46-45fc-8746-76a5b9f30d03" />


---


<img width="1472" height="735" alt="image" src="https://github.com/user-attachments/assets/8c3c654c-29ef-49da-84d3-a5bfa2d37ba0" />




## Tools Used

<p align="center">
  <img src="https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/Power_Query-5B2C91?style=for-the-badge&logo=microsoftpowerpoint&logoColor=white" />
  <img src="https://img.shields.io/badge/Power_Pivot-742774?style=for-the-badge&logo=microsoft&logoColor=white" />
  <img src="https://img.shields.io/badge/DAX-0078D4?style=for-the-badge&logo=microsoftpowerbi&logoColor=white" />
  <img src="https://img.shields.io/badge/VBA-8674A1?style=for-the-badge&logoColor=white" />
  <img src="https://img.shields.io/badge/GitHub-181717?style=for-the-badge&logo=github&logoColor=white" />
</p>


---



---

## Table of Contents

1. [Introduction](#introduction)
2. [Executive Summary](#executive-summary)
3. [Project at a Glance](#1-project-at-a-glance)
4. [Understanding the Business Problem](#2-understanding-the-business-problem)
5. [Data Understanding and Quality Preparation](#3-data-understanding-and-quality-preparation)
6. [Analytical Approach](#4-analytical-approach)
7. [Business Questions Explored](#5-business-questions-explored)
8. [Key Findings and Business Insights](#6-key-findings-and-business-insights)
9. [Business Recommendations & Action Plan](#7-business-recommendations--action-plan)
10. [The Dashboard Solution](#8-the-dashboard-solution)
11. [What This Project Demonstrates](#9-what-this-project-demonstrates)
12. [Important Analytical Considerations](#10-important-analytical-considerations)
13. [Conclusion](#11-conclusion)

---


# Introduction

### Turning customer data into smarter retention decisions

A bank can have 10,000 customers and a substantial customer balance base, yet still face a serious retention problem.

The important question is not simply how many customers are leaving. It is **which customers are leaving, what patterns exist among them, and where should the bank focus its retention efforts?**

That was the question behind this project.

I developed an interactive customer analytics solution for YDP Bank to examine customer churn, understand the characteristics of customers who leave, identify potentially vulnerable segments, and translate the findings into practical retention opportunities.

The goal was to move beyond reporting numbers and use customer data to support better business decisions.


---


## Executive Summary

Customer retention is an important part of sustainable banking growth. However, understanding how many customers leave is only the beginning. Businesses also need to understand which customer segments are most affected and where retention efforts should be prioritised.

This project analyses 10,000 customer records for YDP Bank to investigate customer churn patterns, identify potentially vulnerable customer segments, and uncover opportunities for more targeted retention strategies.

Using Microsoft Excel, Power Pivot, DAX, PivotTables and VBA, I developed an interactive Customer Intelligence Hub featuring two analytical dashboards: **Customer Churn Overview** and **Retention Opportunities**.

The analysis revealed that:

- **20.37% overall churn rate:** 2,037 out of 10,000 customers had exited.
- **Inactive customers showed consistently higher churn:** Across all three tenure groups, inactive customers recorded higher churn rates than active customers.
- **High-balance customers recorded 27.01% churn:** 104 of 385 high-balance customers had exited.
- **Mature Adults recorded the highest displayed age-group churn rate:** 49.45%.
- **Customers holding four products recorded 100% churn:** This unusually high result requires further investigation of the underlying segment size.

The findings highlight customer inactivity, high-balance customer exits, and specific customer segments as areas deserving further investigation.

The project concludes with practical recommendations focused on customer re-engagement, early-tenure support, high-value customer retention, and further investigation of unusual churn patterns.

**The overall objective was to transform raw customer records into actionable business insights that could support more informed customer retention decisions.**

[Download & Explore the Interactive Excel Dashboard](https://github.com/Rofiat-Adebayo/YDP-Bank-Customer-Churn-Retention-Analytics/blob/main/!%20YDP-Bank%20Records%20(version%201).xlsb.xlsm)

*Download the Excel workbook and open it in desktop Excel to explore the interactive dashboards, slicers, and VBA-powered features.*



## 1. Project at a glance

| Project detail | Description |
|---|---|
| Organisation | YDP Bank (project case study) |
| Project | Customer Churn & Retention Analysis |
| Industry | Banking / Financial Services |
| Dataset size | 10,000 customer records |
| Primary objective | Understand customer churn and identify retention opportunities |
| Tools | Microsoft Excel, Power Pivot, DAX, PivotTables, VBA |
| Deliverable | Interactive Excel analytics solution with a homepage and two dashboards |

### The numbers behind the problem

| Key metric | Result |
|---|---:|
| Total customers | 10,000 |
| Customers who exited | 2,037 |
| Customers retained | 7,963 |
| Overall churn rate | 20.37% |
| Retention rate | 79.63% |

**The initial finding:** Approximately one in every five customers had exited.

That raised a more important business question: were these exits distributed evenly across the customer base, or were particular customer segments experiencing higher churn?

---


## Data Source

The dataset used for this project contains 10,000 customer records and includes customer demographics, account information, product ownership, activity status, and customer exit status.

- **Dataset:** YDP Bank Customer Records
- **Records:** 10,000 customer records
- **Format:** Excel (.xlsx)

**Dataset access:** [View and download the dataset](https://github.com/Rofiat-Adebayo/YDP-Bank-Customer-Churn-Retention-Analytics/blob/main/YDP-Bank%20Records%20Dataset%20csv.xlsx)

*Note: YDP Bank is used as a project case study.*

---

## 2. Understanding the business problem

Customer churn is more than a percentage on a dashboard. It represents customers who have stopped using the bank's services.

For a financial institution, losing customers can affect customer relationships, account balances, future product opportunities, and long-term business performance.

However, a single overall churn rate cannot explain where the problem is concentrated.

YDP Bank needed a clearer view of:

- Which customer segments had higher churn rates?
- Were inactive customers more likely to leave?
- Did customer tenure relate to churn?
- Which age groups showed higher churn?
- Did credit score segments show different churn patterns?
- Were customers with higher account balances leaving?
- How did the number of products held relate to customer exits?

These questions shaped the direction of the analysis.

---

## 3. Data understanding and quality preparation

Before building dashboards or calculating KPIs, I first considered the structure and quality of the customer data.

The objective was to make sure the analysis was based on consistent, usable records.

### Data quality observations and actions

| Data quality observation | Action taken |
|---|---|
| Inconsistent or unclear column naming | Standardised column names |
| Inconsistent formatting across records, including customer names or ID formats | Standardised records and formatting |
| Unnecessary row-number column | Removed the unnecessary column |

I also reviewed the available fields to understand what each column represented and how it could contribute to the business questions.

### Understanding the analytical fields

The analysis focused on customer attributes such as:

- Customer ID
- Geography
- Gender
- Age
- Tenure
- Account balance
- Number of products
- Credit card ownership
- Active membership status
- Estimated salary
- Customer exit status

The `Exited` field was particularly important:

- `1` = Customer exited
- `0` = Customer retained

This distinction formed the foundation of the churn calculations.

---

## 4. Analytical approach

Rather than jumping straight into dashboard development, I structured the work around the business questions.

### Step 1: Prepare and organise the data

I reviewed the dataset, addressed the identified quality issues, standardised the relevant fields, and removed unnecessary columns.

### Step 2: Create analytical classifications

I developed customer groupings to make the analysis easier to interpret.

Examples included:

- Age groups
- Credit score bands
- Balance bands
- Tenure bands
- Number of products

These classifications helped transform raw customer attributes into meaningful analytical segments.

### Step 3: Build the analytical model

I used Excel's Data Model and Power Pivot to organise the customer data and create measures using DAX.

The measures supported calculations such as:

- Total customers
- Customers exited
- Customers retained
- Churn rate
- Retention rate
- High-balance customer churn count
- High-balance customer churn rate
- Balance associated with exits
- Churn rate by customer segment

### Step 4: Analyse customer behaviour

Using PivotTables and the calculated measures, I compared churn across customer segments.

I also used conditional formatting to create a churn matrix, making it easier to identify higher-risk combinations of activity status and tenure.

### Step 5: Develop the dashboards

The final analytical solution was organised into two dashboards, supported by a homepage for navigation.

---

## 5. Business questions explored

The analysis was guided by the following questions:

| Business question | Analytical approach |
|---|---|
| What is the overall customer churn situation? | Total customers, exited customers, churn rate and retention rate |
| Are high-balance customers leaving? | High-balance customer churn rate and balance associated with exits |
| Which age groups have higher churn rates? | Churn rate by age group |
| Does geography relate to customer churn? | Churn rate by location |
| How does credit score relate to churn? | Churn rate by credit score band |
| Are inactive customers more likely to leave? | Churn rate by activity status |
| Does tenure relate to customer churn? | Churn rate by tenure band |
| How does the number of products held relate to churn? | Churn rate by number of products |
| Which combinations of activity status and tenure show higher churn? | Activity Status × Tenure matrix |

---

## 6. Key findings and business insights

<img width="787" height="342" alt="image" src="https://github.com/user-attachments/assets/059ae932-5527-45dd-be6b-7cde0ae64287" />


The analysis revealed several patterns that deserve attention.

### Finding 1: Overall customer churn is 20.37%

Out of 10,000 customers, 2,037 had exited, leaving 7,963 retained customers.

This means that approximately 20 out of every 100 customers had exited.

**Business implication:** Customer retention deserves attention as a business priority. However, the overall rate alone does not identify which customers should receive the most attention.

### Finding 2: Female customers account for the larger share of exits

Female customers represented approximately 55.92% of all exited customers.

**Business implication:** This segment accounts for a larger proportion of customer exits. The bank could investigate whether customer needs, service experiences, product preferences, or other factors differ across gender groups.

This finding represents a contribution to total exits, not proof that gender causes churn.

### Finding 3: High-balance customers show a 27.01% churn rate

<img width="816" height="341" alt="image" src="https://github.com/user-attachments/assets/89da7835-9b3e-4276-a4a8-966c82c4ac1d" />


Among the 385 customers classified in the High Balance band, 104 had exited.

That represents a churn rate of 27.01%.

The analysis also identified approximately **₦185.59 million in account balances associated with exited customers**.

**Business implication:** High-balance customer exits deserve further investigation because these customers represent a meaningful portion of the bank's customer balance base.

However, the associated balance should not be interpreted as lost revenue or profit. It represents account balances belonging to customers who exited.

### Finding 4: Mature Adults have the highest displayed age-group churn rate

Customers were grouped into four age categories to make it easier to compare churn patterns across different stages of adulthood:

- **Young Adults:** 18–29 years
- **Adults:** 30–44 years
- **Mature Adults:** 45–59 years
- **Older Adults:** 60 years and above

These age bands provide a consistent basis for comparing customer churn across the different age groups.
The Mature Adults segment recorded a churn rate of 49.45% in the age-group analysis.

Other displayed results included:

| Age group | Churn rate |
|---|---:|
| Adults | 23.67% |
| Mature Adults | 49.45% |
| Older Adults | 27.95% |
| Young Adults | 9.97% |

**Business implication:** The Mature Adults segment warrants further investigation.

The bank could explore whether its products, customer service experience, or engagement approach adequately meet the needs of this segment.

### Finding 5: Inactive customers consistently show higher churn across tenure groups

The Activity Status × Tenure matrix revealed a consistent pattern.

| Activity status | New | Established | Long-Term |
|---|---:|---:|---:|
| Inactive | 27.99% | 26.93% | 25.73% |
| Active | 15.17% | 14.22% | 13.36% |

Inactive customers recorded higher churn rates across all three tenure groups.

The highest displayed rate was among inactive customers in the New tenure group at 27.99%.

**Business implication:** Customer inactivity is an important segment to investigate. The bank could consider targeted engagement, customer check-ins, and relevant product or service interventions before these relationships end.

### Finding 6: Customers holding four products recorded a 100% churn rate

The product analysis showed:

| Number of products | Churn rate |
|---|---:|
| 1 Product | 27.71% |
| 2 Products | 7.58% |
| 3 Products | 82.71% |
| 4 Products | 100.00% |

Customers holding four products recorded the highest displayed churn rate.

**Business implication:** This is an important signal, but it should not be interpreted without checking the number of customers in each product group.

A 100% churn rate can result from a very small group. Before making a retention decision, the bank should examine the actual customer count, exited customer count, and stability of the result.

### Finding 7: The Low credit score band recorded the highest churn rate

| Credit score band | Churn rate |
|---|---:|
| High | 19.32% |
| Low | 22.02% |
| Medium | 20.56% |

The Low credit score band recorded the highest churn rate at 22.02%.

**Business implication:** The bank could investigate whether customer experience, product suitability, or other customer characteristics differ across these segments. Credit score alone should not be treated as an explanation for customer exits.

### Finding 8: Geography also reveals differences in churn

The displayed location analysis showed:

| Location | Churn rate |
|---|---:|
| Enugu | 16.67% |
| Ebonyi | 32.44% |
| Anambra | 16.15% |

Ebonyi recorded the highest displayed churn rate at 32.44%.

**Business implication:** The bank could investigate whether regional differences in customer experience, service accessibility, or engagement contribute to the observed pattern.

These are areas for further investigation, not established causes.

---


## 7. Business Recommendations & Action Plan

The analysis identified several customer segments that deserve closer attention. Based on these findings, the following recommendations outline how YDP Bank could translate the insights into practical retention initiatives.

### 1. Introduce targeted re-engagement campaigns for inactive customers

**Evidence:** Inactive customers recorded higher churn rates than active customers across every tenure group. The highest rate was among inactive customers in the New tenure group at 27.99%.

**Recommended action:**
- Identify inactive customers showing signs of reduced engagement.
- Introduce personalised communication and relevant product offers.
- Review whether customers are experiencing service or accessibility challenges.
- Monitor customer reactivation and subsequent retention.

**Expected business outcome:** Improved customer engagement and potentially fewer exits among inactive customers.

**Suggested KPI:** Customer reactivation rate and churn rate among previously inactive customers.

### 2. Strengthen early-tenure customer engagement

**Evidence:** New-tenure inactive customers recorded the highest churn rate in the Activity Status × Tenure matrix.

**Recommended action:**
- Improve the customer onboarding experience.
- Introduce early-stage customer check-ins.
- Provide guidance on available banking services and products.
- Identify customers who become inactive shortly after joining.

**Expected business outcome:** Stronger early customer relationships and improved retention during the initial customer lifecycle.

**Suggested KPI:** New-customer retention rate and 90-day customer activity rate.

### 3. Prioritise high-balance customer retention

**Evidence:** The High Balance segment recorded a 27.01% churn rate, with 104 of 385 customers exiting. Approximately ₦185.59 million in balances was associated with exited customers.

**Recommended action:**
- Investigate the reasons high-balance customers leave.
- Introduce proactive relationship management for this segment.
- Review customer complaints and service experiences.
- Identify early warning signs of potential customer exits.

**Expected business outcome:** Better protection of valuable customer relationships and the balances associated with them.

**Suggested KPI:** High-balance customer churn rate and balance associated with exited customers.

### 4. Investigate unusually high churn among multi-product customers

**Evidence:** Customers holding four products recorded a 100% churn rate, while the three-product segment recorded 82.71%.

**Recommended action:**
- Validate the number of customers in each product category.
- Investigate whether product complexity, suitability, or service experience may contribute to exits.
- Review product combinations and customer feedback.
- Avoid drawing conclusions from high percentages without considering segment size.

**Expected business outcome:** A clearer understanding of why customers with multiple products are leaving and whether product-related interventions are necessary.

**Suggested KPI:** Churn rate by product count, supported by customer counts and exited customer counts.

### 5. Investigate the Mature Adults customer segment

**Evidence:** Mature Adults recorded the highest displayed age-group churn rate at 49.45%.

**Recommended action:**
- Examine the banking needs and service experiences of this segment.
- Investigate possible differences in product suitability and customer engagement.
- Review customer feedback and exit reasons where available.

**Expected business outcome:** Better understanding of the factors associated with churn within this customer segment.

**Suggested KPI:** Mature Adults churn rate and retention rate.

### 6. Investigate geographical differences in customer churn

**Evidence:** Ebonyi recorded a displayed churn rate of 32.44%, compared with 16.67% in Enugu and 16.15% in Anambra.

**Recommended action:**
- Investigate whether customer service, accessibility, or engagement differs across locations.
- Review customer feedback and complaints by geography.
- Compare customer counts and exit volumes across locations.

**Expected business outcome:** Better visibility into regional retention challenges and opportunities for targeted improvements.

**Suggested KPI:** Churn rate by geography and customer retention by location.


### Recommendation priority

| Priority | Area | Reason |
|---|---|---|
| High | Inactive customers | Higher churn across all tenure groups |
| High | High-balance customers | 27.01% churn rate and significant associated balances |
| High | New-tenure engagement | Highest churn within the activity and tenure matrix |
| Investigate | Three- and four-product customers | Unusually high churn rates; validate segment sizes |
| Investigate | Mature Adults | Highest displayed age-group churn rate |
| Investigate | Geography | Differences in displayed regional churn rates |


### Measuring success

Recommendations should be implemented as measurable initiatives rather than one-off activities.

YDP Bank could establish a retention monitoring process that tracks:

- Overall customer churn rate.
- Customer retention rate.
- Inactive customer reactivation rate.
- High-balance customer churn rate.
- New-customer retention rate.
- Churn rate by product count and geography.

These measures would help the bank evaluate whether its retention initiatives are producing improvements over time.

**Important:** These recommendations are proposed actions based on the observed patterns. The dataset does not establish the underlying causes of customer exits, so further investigation should precede major business interventions.
---


## 8. The dashboard solution

To make the analysis accessible to both technical and non-technical stakeholders, I developed an interactive Excel solution with a dedicated homepage and two analytical dashboards.

### Homepage: YDP Bank Customer Intelligence Hub

The homepage serves as the starting point for the analytical solution.

It provides a clear introduction to the project, a summary of the key customer metrics, and navigation to the two dashboards.

### Dashboard 1: Customer Churn Overview

This dashboard answers the question:

**What does customer churn look like across the bank's customer base?**

It includes:

- Total customers and customers exited
- Overall churn rate
- Retention rate
- Top churn contributor
- Churn rate by activity status and location
- Churn rate by age group
- Churn rate by credit score band

The dashboard provides a high-level understanding of customer exits and the segments where churn is more noticeable.

### Dashboard 2: Retention Opportunities

This dashboard answers the question:

**Where should the bank focus its customer retention efforts?**

It includes:

- Overall churn rate
- High-balance customer churn rate
- Customers exited
- Balance associated with exits
- Churn rate by balance segment
- Activity Status × Tenure matrix
- Churn rate by number of products

The purpose is to move from understanding the churn situation to identifying customer segments that deserve further investigation.

### Interactive features

I also implemented interactive elements to improve how users explore the analysis:

- Slicers for filtering customer segments
- Dynamic PivotTable-based analysis
- Linked pictures to display PivotTable results on dashboards
- Dynamic insight subtitles
- Conditional formatting for the tenure and activity matrix
- VBA-powered filter panel interactions
- Reset View functionality
- Shape-based navigation between worksheets

These features allow users to explore customer segments without manually rebuilding the analysis.

---

## 9. What this project demonstrates

This project demonstrates how I approach a business analytics problem from data preparation through to communicating findings.

It allowed me to apply and develop practical skills in:

**Data preparation**
- Data quality assessment
- Column standardisation
- Record formatting
- Removing unnecessary fields
- Customer segmentation

**Data analysis**
- DAX measures
- Customer churn calculations
- Segment-level analysis
- KPI development
- Comparative analysis

**Business Intelligence**
- Excel Data Model
- Power Pivot
- PivotTables
- Slicers
- Conditional formatting
- Dashboard design
- Dynamic reporting

**Automation and user experience**
- VBA macros
- Interactive filter panels
- Reset functionality
- Worksheet navigation
- Dashboard interactions

Most importantly, it strengthened my ability to connect analytical findings to business questions and communicate what the results could mean for decision-makers.

---

## 10. Important analytical considerations

A few considerations are important when interpreting the results:

1. **Correlation is not causation.** The analysis identifies relationships between customer characteristics and churn, but it does not establish why customers left.

2. **Segment size matters.** Very high churn rates, particularly the 100% rate among four-product customers, should be reviewed alongside the underlying customer count.

3. **Balance is not revenue.** The ₦185.59 million represents balances associated with exited customers. It should not be described as lost revenue or profit.

4. **Further investigation is required.** Additional information about customer complaints, service quality, transaction activity, product usage, and exit reasons would help explain the observed patterns.

5. **The findings are based on the available dataset.** They should not automatically be generalised beyond the customers and period represented in the data.

---

## 11. Conclusion

The most important lesson from this project is that an overall churn rate tells only part of the story.

YDP Bank had a 20.37% customer churn rate, but examining the customer segments revealed more useful patterns.

Inactive customers consistently recorded higher churn across tenure groups. High-balance customers showed a 27.01% churn rate. Mature Adults recorded the highest displayed age-group churn rate, while the four-product segment showed an unusually high result that requires further validation.

These findings provide a starting point for more focused customer retention investigations.

The real value of analytics is not simply knowing what happened. It is understanding where to look next, asking better questions, and giving decision-makers evidence they can use.

**This is what I set out to achieve with the YDP Bank Customer Intelligence Hub: turning customer data into clearer insights and more informed retention decisions.**


```
**Tools used:** Microsoft Excel | Power Pivot | DAX | PivotTables | VBA

**Project:** YDP Bank Customer Churn & Retention Analysis

**Focus:** Customer Analytics | Business Intelligence | Data Storytelling | Retention Strategy
```

<p align="right">
  <a href="#ydp-bank--customer-churn--retention-analysis">⬆️ Back to Top</a>
</p>

