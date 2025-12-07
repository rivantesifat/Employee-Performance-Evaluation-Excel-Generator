## Methodology Overview

This evaluation system is built on a **Weighted Multi-Criteria Decision-Making (MCDM)** framework combined with **statistical normalization** and a **threshold-based penalty model**. The objective is to generate **fair, comparable, and interpretable performance scores** across heterogeneous employee roles.

### 1. Role-Specific Performance Models

Two independent evaluation models are defined:

- **QC Team Model**
- **Production Hand Model**

Each model uses a **custom set of quantitative metrics aligned with the operational responsibilities** of the role. This prevents role-mismatch bias and ensures construct validity of the evaluation.

---

### 2. Statistical Normalization

All positive performance indicators are normalized using **Min–Max Scaling**:

\[
\text{Normalized}(x_i) = \frac{x_i - \min(x)}{\max(x) - \min(x) + \varepsilon}
\]

Where:
- \( x_i \) is the raw value of a metric  
- \( \varepsilon \) is a small constant to prevent division-by-zero

This transforms all metrics into a **unit-free [0, 1] scale**, enabling valid multi-metric aggregation.

---

### 3. Threshold-Based Leave Penalty Model

Leave hours are treated as a **negative performance metric with a tolerance threshold**:

- If `Leave ≤ Threshold` → **No penalty applied**
- If `Leave > Threshold` → **Progressive normalized penalty applied**

\[
\text{LeaveScore}_i = 1 - \frac{\max(0, \text{Leave}_i - T)}{\max(\text{Leave} - T) + \varepsilon}
\]

Where:
- \( T \) is the user-defined leave threshold (default = 50 hours)

This prevents small, acceptable leave usage from distorting performance rankings.

---

### 4. Composite Weighted Scoring

For each employee, the final performance score is computed as:

\[
\text{Final Score} = 20 \times \sum_{j=1}^{n} (w_j \times \text{NormalizedMetric}_j)
\]

Where:
- \( w_j \) are user-defined metric weights  
- \( \sum w_j = 1 \)
- Output scale is standardized to **0–20**

This allows:
- Transparent weighting control
- Real-time scenario testing
- Policy-level score sensitivity analysis

---

### 5. Consistency Modeling (Production Hand Only)

Performance consistency is derived from **historical grade distribution entropy**:

- The **most frequently occurring grade proportion** is used as a consistency index.
- This value is then **min–max normalized** within the Production group.
- Encourages **stable performance behavior**, not short-term spikes.

---

### 6. Transparent Excel-Based Execution Layer

All mathematical logic is **embedded directly as Excel formulas**, enabling:

- Real-time recalculation
- Full auditability
- Zero-code scenario testing for HR and management users

The Python layer acts solely as a **model generator and automation engine**, not as a black-box scorer.

---

### 7. Methodological Strengths

- Role-specific modeling
- Scale-invariant metric fusion
- Non-linear negative penalty control
- Fully interpretable and auditable outputs
- Supports longitudinal and cross-sectional evaluation

This framework is suitable for:
- Workforce performance benchmarking
- Promotion screening
- Operational efficiency diagnostics
- HR analytics research applications
