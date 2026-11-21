# Reproducing the PxPay Transaction Failure Analysis

This document explains the dataset, data preparation, Excel calculations and analysis methods used in the PxPay Transaction Failure Analysis project.

The project was completed using Microsoft Excel. All analytical calculations were formula-driven; PivotTables were not used.

## 1. Project Files

The repository contains:

- `1. Raw Data/pxpay_transactions.csv` – Original synthetic transaction dataset.
- `2. PxPay Transaction Failure Project/PxPay_Failure_Dashboard.xlsx` – Excel workbook containing the dashboard, calculations and final analysis dataset.

## 2. Tools and Requirements

**Software:** Microsoft Excel

**Excel functions used:**

- `COUNTIFS`
- `SUMIFS`
- `IFERROR`
- `INDEX`
- `MATCH`
- `MAX`
- `TEXT`

## 3. Dataset Overview

The original dataset contains **10,030 transaction records and 11 columns**.

The columns include:

| Column | Description |
|---|---|
| `transaction_id` | Unique transaction identifier |
| `customer_id` | Customer identifier |
| `transaction_timestamp` | Date and time of the transaction |
| `transaction_amount_ngn` | Transaction amount in Nigerian Naira |
| `transaction_status` | Successful, Failed or Pending |
| `payment_method` | Payment method used |
| `payment_channel` | Channel through which the payment was initiated |
| `payment_provider` | Payment provider handling the transaction |
| `failure_reason` | Recorded reason for transaction failure |
| `transaction_fee_ngn` | Transaction processing fee |
| `retry_attempt` | Indicates whether a transaction was retried |

The final analysis dataset contains **9,834 records**.

The status distribution is:

| Transaction status | Count |
|---|---:|
| Successful | 8,867 |
| Failed | 704 |
| Pending | 263 |
| **Total** | **9,834** |

## 4. Data Cleaning and Preparation

The original dataset was reviewed and prepared before the final analysis.

The main data-quality checks and treatments included:

### Missing values
- Missing customer IDs were retained and labelled `Unknown User`.
- Missing transaction dates were excluded from time-based analysis.
- Missing transaction amounts and payment metadata were reviewed according to transaction status.
- Missing failure reasons were assigned descriptive labels based on transaction status.

### Inconsistent values
- Transaction statuses were standardised into `Successful`, `Failed` and `Pending`.
- Inconsistent text values and unwanted spaces were reviewed and corrected.
- Data types and date formats were checked.

### Duplicate records
- Duplicate transaction records were investigated.
- Confirmed duplicate copies were removed to avoid inflating transaction counts.

### System anomalies
- Records with missing or invalid information affecting successful transactions were reviewed and separated from the main analysis where appropriate.

**Important:** The original CSV is provided as the raw dataset. The Excel workbook contains the final 9,834-record analysis dataset. It does not contain every intermediate cleaning version or a separate cleaning log, so the cleaning process is documented here rather than presented as a fully automated, one-click reproduction.

## 5. Workbook Structure

The Excel workbook contains five worksheets:

| Worksheet | Purpose |
|---|---|
| `Overview` | Executive dashboard and key performance indicators |
| `Where and Why` | Provider, payment method, channel and failure-reason analysis |
| `Trend and Actions` | Weekly failure trends and business recommendations |
| `Calc` | Supporting calculations and formulas |
| `Data` | Final analysis dataset |

The dashboard includes Provider and Payment Method filters that allow users to explore selected segments of the data.

## 6. KPI Calculations

### Total transactions

Counts all records in the final analysis dataset.

```excel
=COUNTIFS(Data!$D$2:$D$9835,"<>")
```

### Failed transactions

```excel
=COUNTIFS(Data!$D$2:$D$9835,"Failed")
```

### Successful transactions

```excel
=COUNTIFS(Data!$D$2:$D$9835,"Successful")
```

### Pending transactions

```excel
=COUNTIFS(Data!$D$2:$D$9835,"Pending")
```

### Overall failure rate

The failure rate excludes pending transactions because they do not have a final outcome.

**Formula:**

```excel
=IFERROR(Failed/(Failed+Successful),0)
```

Using the project figures:

```excel
=IFERROR(704/(704+8867),0)
```

**Result: 7.36%**

This means approximately 7.36% of completed transactions failed.

## 7. Provider and Payment Method Analysis

The performance of each provider and payment method was compared using failure rates.

### Example: Provider failure count

```excel
=COUNTIFS(Data!$G$2:$G$9835,A6,Data!$D$2:$D$9835,"Failed")
```

Here, `A6` contains the provider name.

### Example: Provider failure rate

```excel
=IFERROR(Failed/(Failed+Successful),0)
```

The same approach was applied to payment methods and channels.

This made it possible to compare performance fairly, rather than relying only on the number of failed transactions.

## 8. Failure Reason Analysis

The number of failed transactions was counted for each recorded failure reason.

### Example: Count a failure reason

```excel
=COUNTIFS(Data!$H$2:$H$9835,A20,Data!$D$2:$D$9835,"Failed")
```

Here, `A20` contains the failure reason.

The individual failure reasons were then grouped into broader categories:

| Category | Failure reasons | Count | Share |
|---|---|---:|---:|
| Connection problems | Timeout, Network Error, USSD Session Expired | 327 | 46% |
| Bank/customer declines | Bank Declined, Insufficient Funds, Issuer Declined | 199 | 28% |
| Provider/system errors | Provider Error, System Error, Wallet Error | 120 | 17% |
| Transfer Pending | Transfer Pending | 48 | 7% |
| Not recorded | Unlogged | 10 | 1% |
| **Total** | | **704** | **100%** |

### Percentage of total failures

```excel
=Failure_Reason_Count/Total_Failed_Transactions
```

This analysis showed that connection problems were the largest group of recorded failure reasons.

## 9. Weekly Trend Analysis

The weekly analysis examined whether the transaction failure rate was increasing, decreasing or remaining relatively stable.

The final analysis dataset includes a transaction week field.

### Failed transactions by week

```excel
=COUNTIFS(Data!$A$2:$A$9835,A34,Data!$D$2:$D$9835,"Failed")
```

Here, `A34` contains the week number.

### Successful transactions by week

```excel
=COUNTIFS(Data!$A$2:$A$9835,A34,Data!$D$2:$D$9835,"Successful")
```

### Weekly failure rate

```excel
=IFERROR(Weekly_Failed/(Weekly_Failed+Weekly_Successful),0)
```

The weekly rates were compared with the overall failure rate to assess changes over the reporting period.

The analysis showed that the failure rate remained relatively stable rather than following a sustained upward trend.

## 10. Failed Transaction Value

The value of failed transactions was calculated using `SUMIFS`.

```excel
=SUMIFS(Data!$C$2:$C$9835,Data!$D$2:$D$9835,"Failed")
```

**Result: approximately ₦12.39 million.**

This represents the recorded value of failed transaction attempts with valid amounts.

It does not represent confirmed lost revenue. Additional information about successful retries, abandoned payments and transaction fee revenue would be required to estimate actual business losses.

## 11. Key Findings

The analysis identified the following:

- Overall transaction failure rate: 7.36%.
- Provider C recorded the highest provider failure rate at 11.4%.
- USSD recorded the highest payment method failure rate at 12.6%.
- Provider C combined with USSD had a failure rate of approximately 17%.
- Connection problems accounted for 46% of all failures.
- Bank declines and insufficient funds accounted for 28%.
- Provider B recorded 44 Provider Error failures.
- Failed transactions represented approximately ₦12.39 million in attempted transaction value.

## 12. Business Recommendations

Based on the findings, the following areas were prioritised for further investigation:

1. Investigate connection-related failures, especially timeouts and network errors.
2. Engage Provider C to understand its high failure rate.
3. Share Provider B's recorded provider errors with its technical team.
4. Review the USSD payment experience and session-expiry issues.
5. Improve failure-reason logging and customer-facing error messages.
6. Monitor the failure rate and measure the impact of corrective actions.

These recommendations are not guaranteed fixes. The analysis identifies patterns and areas requiring attention but does not independently establish the technical root cause of every failure.

## 13. Limitations

- PxPay is fictional, and the dataset is synthetic.
- The analysis identifies patterns but cannot confirm the underlying technical cause of every failure.
- Pending transactions are excluded from the failure-rate calculation.
- Failed attempted transaction value is not confirmed lost revenue.
- The final workbook contains the analysis dataset, not every intermediate cleaning stage.

---

**Purpose of this documentation:** Allow another analyst to understand the project, inspect the calculations and follow the analytical approach used to produce the dashboard and business insights.
