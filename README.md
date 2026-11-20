# PxPay Transaction Failure Analysis

<img width="1280" height="640" alt="pxpay_github_banner" src="https://github.com/user-attachments/assets/d9599e42-f5f6-44b8-8b2a-7dd513a2bfff" />

### Identifying Failure Patterns and Reducing Payment Friction

## Project Overview

| | |
|---|---|
| **Business** | PxPay (Fictional Nigerian Fintech Company) |
| **Period analysed** | Weeks 27–40, 2026 (approximately July–September) |
| **Tool** | Microsoft Excel |
| **Project type** | Business and Data Analytics |
| **Dataset** | Synthetic transaction data |

## The Business Problem

PxPay noticed that many customers were unable to complete their payments.

The business wanted to understand:

- How often are transactions failing?
- Which payment providers and methods have the highest failure rates?
- What are the main reasons for failed transactions?
- What is the financial impact?
- Where should the business focus its improvement efforts?

## What I Did

I analysed 9,834 transaction records using Microsoft Excel.

My work involved:

- Cleaning and standardising transaction data.
- Investigating missing values, duplicate records and data-quality issues.
- Calculating transaction volumes, failure rates and transaction values.
- Comparing payment providers, payment methods and channels.
- Investigating failure reasons and weekly trends.
- Building a formula-driven, filterable Excel dashboard.
- Translating findings into business recommendations.

**All analysis calculations were performed using Excel formulas. No PivotTables were used.**

## Key Findings

| Metric | Result |
|---|---:|
| Original dataset | 10,030 records |
| Final analysis dataset | 9,834 transactions |
| Successful transactions | 8,867 |
| Failed transactions | 704 |
| Pending transactions | 263 |
| Overall failure rate | 7.36% |
| Failed attempted transaction value | ₦12.39 million |

The failure rate excludes pending transactions because they do not yet have a final outcome.

**Failure rate = Failed transactions ÷ (Failed + Successful transactions)**

### Where Do Failures Happen?

- **Provider C** recorded the highest failure rate at 11.4%.
- **USSD** was the weakest payment method at 12.6%.
- Provider C combined with USSD had a failure rate of approximately 17%.
- The API channel recorded a 9.4% failure rate.

### Why Do Failures Happen?

- Connection problems accounted for 46% of all failures.
- Bank declines and insufficient funds accounted for 28%.
- Timeout was the leading individual failure reason, with 165 cases.
- Provider B recorded 44 Provider Error failures.

### Business Impact

Failed transactions represented approximately **₦12.39 million in attempted transaction value**, equivalent to 7.01% of the total recorded transaction value.

This is not confirmed lost revenue. Some customers may have retried successfully or completed their payments through another method.

## Project Preview

The screenshots below provide a quick look at the Excel dashboard and the analysis.

### 1. Executive Dashboard — Overview

A summary of transaction performance, key performance indicators, overall findings and business insights.

<img width="873" height="466" alt="Overview Dashboard" src="https://github.com/user-attachments/assets/14b36a1b-a883-40dd-b732-843305f87973" />

### 2. Where and Why Transactions Fail

A breakdown of payment provider performance, payment methods, channels and the main reasons for failed transactions.

<img width="850" height="461" alt="Where and Why dashboard" src="https://github.com/user-attachments/assets/bb5b9683-a65c-46a2-9149-c15f20762484" />

### 3. Transaction Trends & Business Recommendations

An analysis of weekly transaction failures to identify changes and patterns over the reporting period and the recommended areas of investigation and improvement based on the analysis.

<img width="877" height="461" alt="Trends and Actions Dashboard" src="https://github.com/user-attachments/assets/11904a0a-f959-432d-9c17-4833b4138969" />

## Business Recommendations

Based on the findings, I identified the following priorities:

1. **Investigate connection-related failures:** Examine timeouts, network errors and USSD session expirations.
2. **Engage Provider C:** Investigate why it has the highest failure rate.
3. **Review Provider B errors:** Share the 44 recorded Provider Error cases with its technical team for investigation.
4. **Improve the USSD experience:** Examine session-expiry issues and other connection-related problems.
5. **Improve failure tracking:** Ensure failed transactions have clear and consistent failure reasons.
6. **Monitor performance:** Track the failure rate regularly and measure the impact of corrective actions.

These recommendations are investigation priorities, not guaranteed solutions. The data identifies where attention is needed but does not establish the technical root cause of every failure.

## Tools and Skills Demonstrated

**Tool:** Microsoft Excel

**Technical skills:**
- Data cleaning and preparation
- Data-quality checks
- Excel formulas: `COUNTIFS`, `SUMIFS`, `IFERROR`, `INDEX` and `MATCH`
- KPI calculation
- Failure-rate analysis
- Trend analysis
- Comparative analysis
- Dashboard development

**Business and analytical skills:**
- Business problem definition
- Identifying performance gaps
- Interpreting data and identifying patterns
- Translating findings into business recommendations
- Communicating analytical results to stakeholders

## Repository Contents

| File or folder | Description |
|---|---|
| `1. Raw Data/pxpay_transactions.csv` | Original synthetic transaction dataset |
| `2. PxPay Transactions Failure Project/PxPay_Failure_Dashboard.xlsx` | Excel dashboard, calculations and final analysis dataset |
| `REPRODUCING.md` | Explanation of the analysis and calculation methods |
| `screenshots/` | Dashboard and analysis screenshots |

## Data Disclaimer

PxPay is a fictional Nigerian fintech company. The dataset is synthetic and does not represent actual customer transactions or the operational performance of a real financial institution.

This independent portfolio project was developed to demonstrate practical data analysis, business problem-solving and data storytelling skills.

---

**Project objective:** Identify the main patterns behind transaction failures and recommend evidence-based actions that could help reduce the failure rate.

## If you have any questions or feedback, please reach out directly:

* 💼 **LinkedIn:** [Mowaninuola Umarudeen](https://www.linkedin.com/in/mowaninuolaumarudeen/)
* 📧 **Email:** [mowatheanalyst@gmail.com](mailto:mowatheanalyst@gmail.com)
