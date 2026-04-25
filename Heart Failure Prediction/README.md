# Heart Failure Prediction

A supervised machine learning project that predicts the likelihood of death due to heart failure using real-world clinical data. Built to demonstrate the potential of data-driven approaches in supporting early medical intervention and improving patient outcomes.

---

## Why This Project Matters

Cardiovascular diseases are the leading cause of death worldwide, responsible for an estimated **17.9 million lives lost each year**. In many cases, timely intervention can make the difference between life and death — yet identifying high-risk patients early remains a significant clinical challenge.

This project explores how machine learning can assist medical professionals by flagging patients at elevated risk, enabling faster and more targeted care. While not a replacement for clinical judgment, predictive models like this one can serve as a powerful decision-support tool.

---

## Dataset

- **Source**: [Kaggle — Heart Failure Prediction Dataset](https://www.kaggle.com/code/karnikakapoor/heart-failure-prediction-ann/input)
- **File**: `heart_failure_clinical_records_dataset.csv`
- **Records**: 299 patients
- **Features**: `age`, `anaemia`, `creatinine_phosphokinase`, `diabetes`, `ejection_fraction`, `high_blood_pressure`, `platelets`, `serum_creatinine`, `serum_sodium`, `sex`, `smoking`, `time`
- **Target**: `DEATH_EVENT` — binary classification (0 = survived, 1 = deceased)

---

## Tech Stack

| Category | Tools |
|---|---|
| Language | Python |
| Data Processing | Pandas, NumPy |
| Visualization | Matplotlib, Seaborn |
| Modeling | scikit-learn (SVM, KNN, Logistic Regression) |
| Preprocessing | StandardScaler |

---

## Methodology

### 1. Data Collection & Analysis
- Loaded the dataset and inspected shape, data types, and missing values
- Generated styled descriptive statistics across all features
- Analyzed class distribution — survivors vs. deceased patients
- Computed group means by outcome to highlight feature differences between classes

### 2. Data Visualization
- **Class distribution** bar chart with count labels
- **Feature distribution histograms** for all 7 continuous features split by survival status
- **Correlation heatmap** (lower triangle) showing relationships between all features
- **Age distribution** bar chart split by survival outcome
- **Boxen + swarm plots** for outlier detection across all continuous features (`age`, `creatinine_phosphokinase`, `ejection_fraction`, `platelets`, `serum_creatinine`, `serum_sodium`, `time`)

### 3. Data Preprocessing
- Separated features (`features`) from the target label (`target`)
- Applied **StandardScaler** to normalize all feature values — especially important for SVM and KNN which are sensitive to scale

### 4. Train / Test Split
- Split data into **70% training / 30% test** sets
- `random_state=25` for reproducibility

### 5. Model Training & Evaluation
All three models are evaluated using a shared `evaluate_model()` helper function that prints metrics and renders a **confusion matrix heatmap** for each model.

| Model | Details |
|---|---|
| **Support Vector Machine** | Linear kernel, `random_state=42` |
| **K-Nearest Neighbors** | k=5 neighbours |
| **Logistic Regression** | `random_state=42`, `max_iter=1000` |

Each model is evaluated on: Accuracy, Precision, Recall, F1-Score, Confusion Matrix, and Classification Report.

### 6. Model Comparison
- Results compiled into a styled summary table with color-coded gradient
- Grouped bar chart comparing all four metrics across all three models, with percentage labels on each bar
- Best model identified automatically at runtime

---

## Model Performance

| Model | Accuracy | Precision | Recall | F1-Score |
|---|---|---|---|---|
| Support Vector Machine | ~82% | — | — | — |
| K-Nearest Neighbors | ~78% | — | — | — |
| **Logistic Regression** | **~84%** | — | — | — |

> Exact metric values are printed and visualized at runtime.

**Logistic Regression** achieves the highest accuracy and is identified automatically as the best-performing model. Its strong interpretability makes it especially suitable for clinical decision-support contexts.

---

## Future Improvements

- **Ensemble methods** — Experiment with Random Forest, XGBoost, or stacking classifiers to push accuracy further.
- **Feature importance analysis** — Identify which clinical indicators are the strongest predictors of mortality.
- **Hyperparameter tuning** — Apply `GridSearchCV` or Bayesian optimization for finer model performance.
- **Class imbalance handling** — Explore SMOTE or class weighting, as the dataset has more survivors than deceased patients, which can bias model predictions.
- **Deployment** — Wrap the best model in a Streamlit or Flask web interface for interactive, real-time predictions.

---

## How to Run

1. Clone the repository or download the project folder.
2. Install the required dependencies:
   ```bash
   pip install pandas numpy matplotlib seaborn scikit-learn
   ```
3. Place `heart_failure_clinical_records_dataset.csv` in the same directory as the notebook.
4. Open `heart_failure_PREDICTION.ipynb` in Jupyter Notebook and run all cells.

---

## License

This project is intended for educational and portfolio purposes.
