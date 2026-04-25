# Diabetes Prediction using Machine Learning

A supervised machine learning project that predicts whether a patient has diabetes based on key medical indicators. Built using a Support Vector Machine classifier trained on the PIMA Diabetes dataset, this project demonstrates a full ML pipeline — from data exploration and visualization through to model serialization and a live predictive function.

---

## Why This Project Matters

Diabetes affects over **500 million people worldwide** and remains one of the most prevalent chronic diseases globally. Early and accurate detection is critical — undiagnosed diabetes can lead to severe complications including heart disease, kidney failure, and vision loss.

By leveraging machine learning on clinical data, this project shows how predictive models can support faster, data-driven screening, potentially helping healthcare providers identify at-risk patients before symptoms escalate.

---

## Project Structure

```
├── diabetes-prediction-with-machine-learning.ipynb  # Full ML pipeline notebook
├── diabetes.csv                                      # PIMA Diabetes dataset
└── diabetes_prediction_model.pkl                     # Serialized trained model
```

---

## Dataset

The project uses the **PIMA Indians Diabetes Dataset**, originally provided by the National Institute of Diabetes and Digestive and Kidney Diseases.

| Feature | Description |
|---|---|
| `Pregnancies` | Number of pregnancies |
| `Glucose` | Plasma glucose concentration |
| `BloodPressure` | Diastolic blood pressure (mm Hg) |
| `SkinThickness` | Triceps skinfold thickness (mm) |
| `Insulin` | 2-hour serum insulin (μU/ml) |
| `BMI` | Body mass index |
| `DiabetesPedigreeFunction` | Diabetes likelihood based on family history |
| `Age` | Age in years |
| `Outcome` | Target — 0: Non-Diabetic, 1: Diabetic |

---

## Tech Stack

| Category | Tools |
|---|---|
| Language | Python |
| Data Processing | Pandas, NumPy |
| Visualization | Matplotlib, Seaborn |
| Preprocessing | scikit-learn (StandardScaler) |
| Modeling | scikit-learn (SVM — linear kernel) |
| Model Serialization | pickle |

---

## Methodology

### 1. Data Collection & Analysis
- Loaded the PIMA dataset and inspected its shape, data types, and missing values
- Generated styled descriptive statistics across all features
- Analyzed class distribution — 500 non-diabetic vs. 268 diabetic patients
- Computed group means by outcome to highlight feature differences between classes

### 2. Data Visualization
- **Class distribution** bar chart with count labels
- **Feature distribution histograms** for 6 key features (Glucose, BMI, Age, BloodPressure, Insulin, SkinThickness) split by diabetes status
- **Correlation heatmap** (lower triangle) showing relationships between all features
- **Boxplots** comparing Glucose and BMI distributions across diabetic and non-diabetic groups

### 3. Data Preprocessing
- Separated features (`features`) from the target label (`target`)
- Applied **StandardScaler** to normalize all features to zero mean and unit variance — essential for SVM performance

### 4. Train / Test Split
- Split data into **80% training / 20% test** sets
- Used **stratified sampling** (`stratify=target`) to preserve class ratios in both splits
- `random_state=2` for reproducibility

### 5. Model Training
- Trained a **Support Vector Machine** (`svm_model`) with a linear kernel
- Prints the number of support vectors per class after training

### 6. Model Evaluation
- Evaluated accuracy on both **training set** (78.66%) and **test set** (77.27%)
- Generated a full **classification report** with precision, recall, and F1-score
- Visualized a **confusion matrix heatmap** with a breakdown of TP, TN, FP, and FN counts

### 7. Predictive Function
- Implemented a clean `predict_diabetes()` function that takes a tuple of patient measurements, standardizes them, and returns a labeled prediction result
- Example input: `(5, 166, 72, 19, 175, 25.8, 0.587, 51)` → **Diabetic**

### 8. Model Serialization
- Saved the trained model to `diabetes_prediction_model.pkl` using `pickle`
- Verified the saved model reloads correctly and achieves the same test accuracy

---

## Model Performance

| Dataset | Accuracy |
|---|---|
| Training Set | 78.66% |
| Test Set | 77.27% |

The SVM linear kernel strikes a strong balance between simplicity and performance, making the model well-suited for a medical classification task where interpretability and reliability matter.

---

## Future Improvements

- **Hyperparameter tuning** — Use `GridSearchCV` to find the optimal SVM kernel and regularization parameters.
- **Additional classifiers** — Benchmark against Logistic Regression, Random Forest, or XGBoost for comparison.
- **Feature importance analysis** — Investigate which features (e.g., Glucose, BMI) contribute most to predictions.
- **Class imbalance handling** — Apply SMOTE or class weighting to address the imbalance between diabetic and non-diabetic samples.
- **Web deployment** — Integrate the serialized model into a Flask or Streamlit app for real-time predictions via a user interface.

---

## How to Run

1. Clone the repository or download the project folder.
2. Install the required dependencies:
   ```bash
   pip install numpy pandas matplotlib seaborn scikit-learn
   ```
3. Place `diabetes.csv` in the same directory as the notebook.
4. Open `diabetes-prediction-with-machine-learning.ipynb` in Jupyter Notebook and run all cells.

To load and use the saved model directly:
```python
import pickle
with open("diabetes_prediction_model.pkl", "rb") as model_file:
    loaded_model = pickle.load(model_file)
```

---

## License

This project is licensed under the MIT License.

## Acknowledgements

- Dataset provided by the **National Institute of Diabetes and Digestive and Kidney Diseases**
- Powered by Python and scikit-learn
