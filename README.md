# Statistical Employee Performance Evaluation System (Python → Excel)

This repository contains a **Python-based statistical modeling system** that generates a **fully dynamic, auditable, and customizable Excel-based employee performance evaluation framework**.  

The system is designed to support **objective, role-specific, multi-criteria performance assessment** for two distinct operational groups:

- **Quality Control (QC) Team**
- **Production Hand Team**

The output is a **decision-ready Excel workbook** where:
- All scores are computed using transparent formulas
- All metric weights are user-adjustable
- Leave penalties are threshold-controlled
- Final scores are normalized and ranked automatically

This framework is suitable for:
- HR analytics
- Promotion screening
- Workforce benchmarking
- Operational efficiency studies
- Applied statistical modeling research

---

## 1. Conceptual Framework

The system is based on a **Weighted Multi-Criteria Decision-Making (MCDM)** approach combined with:

- Statistical **min–max normalization**
- **Threshold-based negative penalty functions**
- **Composite linear aggregation**
- **Role-specific metric selection**

Each employee is evaluated using only the metrics that are **operationally meaningful** for their specific job category. This avoids structural bias caused by cross-role metric contamination.

---

## 2. Role-Specific Evaluation Structure

### 2.1 QC Team Metrics

Each QC employee is evaluated using:

| Metric | Type |
|--------|------|
| Total Work Hours | Positive |
| Leave Hours | Negative (Threshold-controlled) |
| Distinct Tasks | Positive |
| Projects Worked | Positive |
| Total Tasks | Positive |
| Team Lead Score | Positive |

### 2.2 Production Hand Metrics

Each Production employee is evaluated using:

| Metric | Type |
|--------|------|
| Production Hours | Positive |
| Leave Hours | Negative (Threshold-controlled) |
| Production-to-QC Ratio | Positive |
| Distinct Tasks | Positive |
| Projects Worked | Positive |
| Total Tasks | Positive |
| Average Evaluation (out of 10) | Positive |
| Consistency Index | Positive |

---

## 3. Data Normalization Model

All positive performance indicators are normalized using **min–max scaling**:

\[
N(x_i) = \frac{x_i - \min(x)}{\max(x) - \min(x) + \varepsilon}
\]

Where:

- \( x_i \) = raw metric value of employee *i*
- \( \varepsilon \) = small constant to avoid division-by-zero
- Output range: \( [0,1] \)

This removes unit dependency (hours, counts, ratios) and enables valid metric fusion.

---

## 4. Threshold-Based Leave Penalty Model

Leave is modeled as a **negative indicator with tolerance control**:

Let:
- \( L_i \) = leave hours of employee *i*
- \( T \) = user-defined threshold (default = 50 hours)

Then:

\[
\text{LeaveScore}_i =
\begin{cases}
1, & \text{if } L_i \le T \\\\
1 - \frac{L_i - T}{\max(L - T) + \varepsilon}, & \text{if } L_i > T
\end{cases}
\]

This ensures:
- No penalty for acceptable leave usage
- Progressive penalty only for excessive absence
- Group-relative fairness in normalization

---

## 5. Consistency Modeling (Production Team)

Consistency is derived from **historical performance grades**:

Let:
- \( n \) = total number of recorded grades
- \( f_{\max} \) = frequency of the most common grade

Then raw consistency is defined as:

\[
C = \frac{f_{\max}}{n}
\]

This captures **behavioral stability**, not just performance magnitude.

This value is then normalized using min–max scaling and treated as a positive metric.

---

## 6. Composite Final Score Model

For each employee, the final score is computed as:

\[
\text{Final Score} = 20 \times \sum_{j=1}^{m} \left( w_j \cdot N_j \right)
\]

Where:
- \( N_j \) = normalized value of metric *j*
- \( w_j \) = user-defined metric weight
- \( \sum w_j = 1 \)
- Output scale: **0–20**

This allows:

- Transparent trade-off control
- Policy-based weighting
- Sensitivity testing
- Reproducible ranking

---

## 7. Excel Execution Layer

All scoring logic is embedded as **native Excel formulas**, including:

- Normalization formulas
- Leave penalty equations
- Weighted aggregation using `SUMPRODUCT`
- Automatic ranking functions

This ensures:

- No hidden black-box logic
- Real-time recalculation
- Audit-ready model behavior
- Zero-code usage for HR users

The **Python layer acts purely as the model generator**, not as the execution engine.

---

## 8. Python Automation Layer

The Python script performs:

- Structured data ingestion
- Derived metric computation
- Consistency mapping
- Excel workbook construction
- Formula injection
- Styling and layout automation

Key dependencies:
- `pandas`
- `numpy`
- `openpyxl`

---

## 9. Customization Workflow

Users can customize the system in three ways:

1. **Edit employee data** inside the Python script
2. **Adjust metric weights directly inside Excel**
3. **Modify the leave threshold inside Excel**

All modifications immediately propagate to:

- Normalized metrics
- Final scores
- Rankings

---

## 10. Methodological Strengths

- Role-specific construct validity
- Scale-invariant multi-metric fusion
- Non-linear negative penalty control
- Fully interpretable Excel execution
- Reproducible statistical outputs
- Zero black-box dependencies

---

## 11. Limitations

- Linear weighted aggregation assumes metric independence
- Consistency is grade-frequency based (not variance-based)
- No stochastic uncertainty modeling (deterministic framework)

---

## 12. Application Domains

- HR performance evaluation
- Workforce analytics
- Promotion benchmarking
- Operational efficiency diagnostics
- Applied statistics research
- Decision-support systems

---

## Author

**Md. Nashid Kamal Sifat**  
Geo-ICT Engineer | Data Science & Applied Statistics  
Email: nashidsifat@outlook.com
