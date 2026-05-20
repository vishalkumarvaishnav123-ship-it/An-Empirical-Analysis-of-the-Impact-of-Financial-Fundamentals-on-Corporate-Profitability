# Thesis: An Empirical Analysis of the Impact of Financial Fundamentals on Corporate Profitability

**A Comparison across Countries (Germany vs United Kingdom)**

## Overview

This repository contains the data, analysis code, and LaTeX source files for the thesis. The study examines how three financial fundamentals — Equity/Total Assets, Quick Ratio, and Sales/Total Assets — affect corporate profitability (ROA and ROE) for 35 German and 35 UK firms in 2021 and 2024.

## Repository Structure

```
├── data/
│   └── German_compnies_data_corrected_with_uk_and_germany___Kopie_00.xlsx
├── figures/                          # Generated figures (output)
│   ├── Figure_2_1_Conceptual_Framework.png
│   ├── Figure_3_1_Sample_Distribution.png
│   ├── Figure_3_2_Mean_ROA_ROE.png
│   ├── Figure_4_1_Scatter_ROA.png
│   ├── Figure_4_2_Scatter_ROE.png
│   ├── Figure_4_3_R_Squared.png
│   ├── Figure_5_1_Coefficient_Comparison.png
│   └── Figure_5_2_Hypotheses_Outcomes.png
├── latex/
│   ├── Chapter_1_Introduction.tex
│   ├── Chapter_2_Literature_Review.tex
│   ├── Chapter_3_Data_Methodology.tex
│   ├── Chapter_4_Empirical_Results.tex
│   ├── Chapter_5_Discussion.tex
│   └── Chapter_6_Conclusion.tex
├── generate_figures.py               # Python script to generate all 8 figures
├── requirements.txt
└── README.md
```

## How to Run

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Generate all figures

```bash
python generate_figures.py
```

This reads the Excel data file and generates all 8 figures in the `figures/` directory.

### 3. Compile LaTeX

```bash
cd latex
pdflatex Chapter_1_Introduction.tex
pdflatex Chapter_2_Literature_Review.tex
pdflatex Chapter_3_Data_Methodology.tex
pdflatex Chapter_4_Empirical_Results.tex
pdflatex Chapter_5_Discussion.tex
pdflatex Chapter_6_Conclusion.tex
```

> **Note:** Copy all PNG files from `figures/` into the `latex/` directory before compiling.

## Methodology

- **Method:** OLS Regression with a two-step approach
- **Dependent variables:** ROA, ROE
- **Independent variables:** Equity/TA (+), Quick Ratio (+), Sales/TA (+)
- **Samples:** Germany 2021, Germany 2024, UK 2021, UK 2024
- **Total models:** 8 (2 dependent × 4 samples)

## Key Findings

| Variable | Significant in | Direction |
|---|---|---|
| Equity/TA | 5 of 8 models | Positive (2021), Negative (UK 2024 ROE) |
| Quick Ratio | 0 of 8 models | — |
| Sales/TA | 3 of 8 models | Negative (Germany only) |

## Tools Used

- **Python 3.10+** — Data analysis and figure generation
- **statsmodels** — OLS regression
- **matplotlib** — Figures and charts
- **pandas / openpyxl** — Data loading and processing
- **LaTeX** — Thesis document preparation
