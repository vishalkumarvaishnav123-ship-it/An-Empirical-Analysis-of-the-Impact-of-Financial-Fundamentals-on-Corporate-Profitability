"""
Thesis: An Empirical Analysis of the Impact of Financial Fundamentals on Corporate Profitability
Author: Vishal
Description: This script generates all 8 figures used in the thesis.
Requirements: pip install pandas numpy matplotlib statsmodels openpyxl
"""

import pandas as pd
import numpy as np
import statsmodels.api as sm
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os

# ============================================
# CONFIGURATION
# ============================================
DATA_FILE = 'German_compnies_data_corrected_with_uk_and_germany___Kopie_00.xlsx'
OUTPUT_DIR = 'figures'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Sheet mapping
SHEET_MAP = {
    'Germany 2021': 'german2021',
    'Germany 2024': 'german 2024',
    'UK 2021': 'UK2021',
    'UK 2024': 'UK2024'
}

# Color palette
COLORS = {
    'Germany 2021': '#4472C4',
    'Germany 2024': '#5B9BD5',
    'UK 2021': '#ED7D31',
    'UK 2024': '#F4B183'
}

BLUE = '#4472C4'
ORANGE = '#ED7D31'


# ============================================
# DATA LOADING
# ============================================
def load_clean(sheet_name):
    """Load and clean a single sheet from the Excel file."""
    df = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
    col_map = {}
    for c in df.columns:
        cl = c.strip().lower()
        if 'firm' in cl:
            col_map[c] = 'Firm'
        elif 'euity' in cl or 'equity_ta' in cl:
            col_map[c] = 'Equity_TA'
        elif 'quick' in cl:
            col_map[c] = 'Quick_Ratio'
        elif 'sal' in cl and 'ta' in cl:
            col_map[c] = 'Sales_TA'
        elif 'roa' in cl:
            col_map[c] = 'ROA'
        elif 'roe' in cl:
            col_map[c] = 'ROE'
    df = df.rename(columns=col_map)
    needed = ['Firm', 'Equity_TA', 'Quick_Ratio', 'Sales_TA', 'ROA', 'ROE']
    df = df[needed].dropna(subset=['Firm'])
    for c in needed[1:]:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    df = df.dropna()
    return df.reset_index(drop=True)


def load_all_data():
    """Load all four datasets."""
    datasets = {}
    for label, sheet in SHEET_MAP.items():
        datasets[label] = load_clean(sheet)
        print(f"Loaded {label}: {len(datasets[label])} firms")
    return datasets


def run_regressions(datasets):
    """Run OLS regressions for all 8 models."""
    results = {}
    data_labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    for label in data_labels:
        df = datasets[label]
        for dep in ['ROA', 'ROE']:
            X = sm.add_constant(df[['Equity_TA', 'Quick_Ratio', 'Sales_TA']])
            model = sm.OLS(df[dep], X).fit()
            results[f"{label} - {dep}"] = model
    return results


# ============================================
# FIGURE 2.1: Conceptual Framework
# ============================================
def figure_2_1():
    """Generate Figure 2.1: Conceptual Framework."""
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Independent variables
    indep_vars = [
        'Equity / Total Assets\n(Capital Structure)',
        'Quick Ratio\n(Liquidity)',
        'Sales / Total Assets\n(Asset Efficiency)'
    ]
    for i, var in enumerate(indep_vars):
        y = 4.5 - i * 1.5
        ax.add_patch(plt.Rectangle(
            (0.5, y - 0.4), 3.5, 0.8,
            fill=True, facecolor=BLUE, edgecolor='black', linewidth=1.5
        ))
        ax.text(2.25, y, var, ha='center', va='center',
                fontsize=10, color='white', fontweight='bold')
        ax.annotate('', xy=(5.5, 3), xytext=(4, y),
                    arrowprops=dict(arrowstyle='->', color='#333333', lw=2))
        ax.text(4.6, y + 0.15, '(+)', fontsize=9, color='green', fontweight='bold')

    # Dependent variables
    dep_vars = ['ROA\n(Return on Assets)', 'ROE\n(Return on Equity)']
    for i, var in enumerate(dep_vars):
        y = 3.7 - i * 1.4
        ax.add_patch(plt.Rectangle(
            (5.5, y - 0.4), 3.5, 0.8,
            fill=True, facecolor=ORANGE, edgecolor='black', linewidth=1.5
        ))
        ax.text(7.25, y, var, ha='center', va='center',
                fontsize=10, color='white', fontweight='bold')

    # Labels
    ax.text(2.25, 5.5, 'Independent Variables',
            ha='center', va='center', fontsize=12, fontweight='bold', color=BLUE)
    ax.text(7.25, 5.5, 'Dependent Variables',
            ha='center', va='center', fontsize=12, fontweight='bold', color=ORANGE)

    fig.suptitle('Figure 2.1: Conceptual Framework', fontsize=14, fontweight='bold', y=0.98)
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, 'Figure_2_1_Conceptual_Framework.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 3.1: Sample Distribution
# ============================================
def figure_3_1(datasets):
    """Generate Figure 3.1: Sample Distribution by Country and Year."""
    fig, ax = plt.subplots(figsize=(8, 5))
    categories = ['Germany\n2021', 'Germany\n2024', 'UK\n2021', 'UK\n2024']
    labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    counts = [len(datasets[k]) for k in labels]
    colors = [BLUE, BLUE, ORANGE, ORANGE]
    hatches = ['', '//', '', '//']

    bars = ax.bar(categories, counts, color=colors, edgecolor='black', linewidth=1)
    for bar, h in zip(bars, hatches):
        bar.set_hatch(h)
    for bar, c in zip(bars, counts):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                str(c), ha='center', va='bottom', fontsize=12, fontweight='bold')

    ax.set_ylabel('Number of Firms', fontsize=12)
    ax.set_title('Figure 3.1: Sample Distribution by Country and Year',
                 fontsize=13, fontweight='bold')
    ax.set_ylim(0, max(counts) + 5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, 'Figure_3_1_Sample_Distribution.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 3.2: Mean ROA and ROE
# ============================================
def figure_3_2(datasets):
    """Generate Figure 3.2: Bar Chart of Mean ROA and ROE across Countries."""
    fig, ax = plt.subplots(figsize=(9, 5))
    labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    mean_roa = [datasets[k]['ROA'].mean() for k in labels]
    mean_roe = [datasets[k]['ROE'].mean() for k in labels]
    x = np.arange(len(labels))
    w = 0.35

    bars1 = ax.bar(x - w / 2, mean_roa, w, label='Mean ROA', color=BLUE, edgecolor='black')
    bars2 = ax.bar(x + w / 2, mean_roe, w, label='Mean ROE', color=ORANGE, edgecolor='black')

    for bar in bars1:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.003,
                f'{bar.get_height():.3f}', ha='center', va='bottom', fontsize=9)
    for bar in bars2:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.003,
                f'{bar.get_height():.3f}', ha='center', va='bottom', fontsize=9)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel('Mean Value', fontsize=12)
    ax.set_title('Figure 3.2: Mean ROA and ROE across Countries and Years',
                 fontsize=13, fontweight='bold')
    ax.legend()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, 'Figure_3_2_Mean_ROA_ROE.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 4.1: Scatter Plots – IV vs ROA
# ============================================
def figure_4_1(datasets):
    """Generate Figure 4.1: Scatter Plots – Independent Variables vs ROA."""
    indep = ['Equity_TA', 'Quick_Ratio', 'Sales_TA']
    indep_labels = ['Equity/TA', 'Quick Ratio', 'Sales/TA']
    data_labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    colors_list = list(COLORS.values())

    fig, axes = plt.subplots(3, 4, figsize=(18, 12))
    for row, (var, vlabel) in enumerate(zip(indep, indep_labels)):
        for col, (dlabel, clr) in enumerate(zip(data_labels, colors_list)):
            ax = axes[row, col]
            df = datasets[dlabel]
            ax.scatter(df[var], df['ROA'], color=clr, alpha=0.7,
                       edgecolors='black', s=50)
            # Trend line
            z = np.polyfit(df[var], df['ROA'], 1)
            p = np.poly1d(z)
            xline = np.linspace(df[var].min(), df[var].max(), 100)
            ax.plot(xline, p(xline), '--', color='red', linewidth=1.5)
            ax.set_xlabel(vlabel, fontsize=9)
            ax.set_ylabel('ROA', fontsize=9)
            ax.set_title(f'{dlabel}', fontsize=10, fontweight='bold')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

    fig.suptitle('Figure 4.1: Scatter Plots – Independent Variables vs ROA',
                 fontsize=14, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    path = os.path.join(OUTPUT_DIR, 'Figure_4_1_Scatter_ROA.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 4.2: Scatter Plots – IV vs ROE
# ============================================
def figure_4_2(datasets):
    """Generate Figure 4.2: Scatter Plots – Independent Variables vs ROE."""
    indep = ['Equity_TA', 'Quick_Ratio', 'Sales_TA']
    indep_labels = ['Equity/TA', 'Quick Ratio', 'Sales/TA']
    data_labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    colors_list = list(COLORS.values())

    fig, axes = plt.subplots(3, 4, figsize=(18, 12))
    for row, (var, vlabel) in enumerate(zip(indep, indep_labels)):
        for col, (dlabel, clr) in enumerate(zip(data_labels, colors_list)):
            ax = axes[row, col]
            df = datasets[dlabel]
            ax.scatter(df[var], df['ROE'], color=clr, alpha=0.7,
                       edgecolors='black', s=50)
            z = np.polyfit(df[var], df['ROE'], 1)
            p = np.poly1d(z)
            xline = np.linspace(df[var].min(), df[var].max(), 100)
            ax.plot(xline, p(xline), '--', color='red', linewidth=1.5)
            ax.set_xlabel(vlabel, fontsize=9)
            ax.set_ylabel('ROE', fontsize=9)
            ax.set_title(f'{dlabel}', fontsize=10, fontweight='bold')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

    fig.suptitle('Figure 4.2: Scatter Plots – Independent Variables vs ROE',
                 fontsize=14, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    path = os.path.join(OUTPUT_DIR, 'Figure_4_2_Scatter_ROE.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 4.3: R-Squared Comparison
# ============================================
def figure_4_3(reg_results):
    """Generate Figure 4.3: R-Squared and Adjusted R-Squared across All Models."""
    fig, ax = plt.subplots(figsize=(10, 5))
    model_names = list(reg_results.keys())
    r2_vals = [reg_results[k].rsquared for k in model_names]
    adj_r2_vals = [reg_results[k].rsquared_adj for k in model_names]
    x = np.arange(len(model_names))
    w = 0.35

    bars1 = ax.bar(x - w / 2, r2_vals, w, label='R²', color=BLUE, edgecolor='black')
    bars2 = ax.bar(x + w / 2, adj_r2_vals, w, label='Adjusted R²',
                   color=ORANGE, edgecolor='black')

    for bar in bars1:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.005,
                f'{bar.get_height():.3f}', ha='center', va='bottom', fontsize=8)
    for bar in bars2:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.005,
                f'{bar.get_height():.3f}', ha='center', va='bottom', fontsize=8)

    ax.set_xticks(x)
    ax.set_xticklabels([k.replace(' - ', '\n') for k in model_names], fontsize=8)
    ax.set_ylabel('Value', fontsize=11)
    ax.set_title('Figure 4.3: R-Squared and Adjusted R-Squared across All Models',
                 fontsize=13, fontweight='bold')
    ax.legend()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, 'Figure_4_3_R_Squared.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 5.1: Coefficient Comparison
# ============================================
def figure_5_1(reg_results):
    """Generate Figure 5.1: Comparison of Significant Coefficients (Germany vs UK)."""
    data_labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    vars_list = ['Equity_TA', 'Quick_Ratio', 'Sales_TA']
    var_labels = ['Equity/TA', 'Quick Ratio', 'Sales/TA']
    colors_list = list(COLORS.values())

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    for idx, dep in enumerate(['ROA', 'ROE']):
        ax = axes[idx]
        x = np.arange(len(vars_list))
        w = 0.18
        for i, (label, clr) in enumerate(zip(data_labels, colors_list)):
            key = f"{label} - {dep}"
            model = reg_results[key]
            coefs = [model.params[v] for v in vars_list]
            pvals = [model.pvalues[v] for v in vars_list]
            bars = ax.bar(x + i * w - 1.5 * w, coefs, w, label=label,
                          color=clr, edgecolor='black')
            for j, (bar, pv) in enumerate(zip(bars, pvals)):
                if pv <= 0.1:
                    sig = '***' if pv <= 0.01 else ('**' if pv <= 0.05 else '*')
                    y_pos = (bar.get_height() + 0.01 if bar.get_height() >= 0
                             else bar.get_height() - 0.04)
                    ax.text(bar.get_x() + bar.get_width() / 2, y_pos, sig,
                            ha='center', fontsize=9, fontweight='bold', color='red')
        ax.set_xticks(x)
        ax.set_xticklabels(var_labels, fontsize=10)
        ax.set_ylabel('Coefficient', fontsize=11)
        ax.set_title(f'Dependent Variable: {dep}', fontsize=12, fontweight='bold')
        ax.legend(fontsize=8)
        ax.axhline(y=0, color='black', linewidth=0.5)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    fig.suptitle(
        'Figure 5.1: Comparison of Coefficients across Countries '
        '(* p<0.10, ** p<0.05, *** p<0.01)',
        fontsize=13, fontweight='bold'
    )
    plt.tight_layout(rect=[0, 0, 1, 0.93])
    path = os.path.join(OUTPUT_DIR, 'Figure_5_1_Coefficient_Comparison.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# FIGURE 5.2: Hypotheses Outcomes Summary
# ============================================
def figure_5_2(reg_results):
    """Generate Figure 5.2: Hypotheses Outcomes Summary (Visual Grid)."""
    data_labels = ['Germany 2021', 'Germany 2024', 'UK 2021', 'UK 2024']
    vars_list = ['Equity_TA', 'Quick_Ratio', 'Sales_TA']
    var_labels = ['Equity/TA', 'Quick Ratio', 'Sales/TA']

    outcome_colors = {
        'Supported': '#70AD47',
        'Rejected': '#FF0000',
        'Not Significant': '#BFBFBF'
    }
    outcome_symbols = {
        'Supported': '✓',
        'Rejected': '✗',
        'Not Significant': '—'
    }

    # Build grid
    rows = []
    for dep in ['ROA', 'ROE']:
        for vlabel in var_labels:
            rows.append(f"{dep} ~ {vlabel}")

    cell_text = []
    cell_colors = []
    for dep in ['ROA', 'ROE']:
        for var, vlabel in zip(vars_list, var_labels):
            row_text = []
            row_colors = []
            for label in data_labels:
                key = f"{label} - {dep}"
                model = reg_results[key]
                coef = model.params[var]
                pval = model.pvalues[var]
                sig = pval <= 0.10
                actual = '+' if coef > 0 else '-'
                expected = '+'
                if sig and actual == expected:
                    outcome = 'Supported'
                elif sig and actual != expected:
                    outcome = 'Rejected'
                else:
                    outcome = 'Not Significant'
                row_text.append(f"{outcome_symbols[outcome]} {outcome}")
                row_colors.append(outcome_colors[outcome])
            cell_text.append(row_text)
            cell_colors.append(row_colors)

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.axis('off')

    table = ax.table(
        cellText=cell_text, rowLabels=rows, colLabels=data_labels,
        loc='center', cellLoc='center'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.8)

    # Color cells
    for i, row_c in enumerate(cell_colors):
        for j, clr in enumerate(row_c):
            table[i + 1, j].set_facecolor(clr + '40')

    # Style headers
    for j in range(len(data_labels)):
        table[0, j].set_facecolor(BLUE)
        table[0, j].set_text_props(color='white', fontweight='bold')
    for i in range(len(rows)):
        table[i + 1, -1].set_facecolor('#F2F2F2')
        table[i + 1, -1].set_text_props(fontweight='bold')

    ax.set_title(
        'Figure 5.2: Hypotheses Outcomes Summary\n(Expected sign: + for all)',
        fontsize=13, fontweight='bold', pad=20
    )
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, 'Figure_5_2_Hypotheses_Outcomes.png')
    plt.savefig(path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"Saved: {path}")


# ============================================
# MAIN
# ============================================
def main():
    print("=" * 60)
    print("Thesis Figure Generator")
    print("=" * 60)

    # Load data
    print("\n[1/3] Loading data...")
    datasets = load_all_data()

    # Run regressions
    print("\n[2/3] Running regressions...")
    reg_results = run_regressions(datasets)

    # Print regression summary
    print("\n--- Regression Summary ---")
    for key, model in reg_results.items():
        sig_vars = [v for v in ['Equity_TA', 'Quick_Ratio', 'Sales_TA']
                    if model.pvalues[v] <= 0.10]
        print(f"{key}: R²={model.rsquared:.4f}, "
              f"Adj R²={model.rsquared_adj:.4f}, "
              f"Significant: {sig_vars if sig_vars else 'None'}")

    # Generate all figures
    print("\n[3/3] Generating figures...")
    figure_2_1()
    figure_3_1(datasets)
    figure_3_2(datasets)
    figure_4_1(datasets)
    figure_4_2(datasets)
    figure_4_3(reg_results)
    figure_5_1(reg_results)
    figure_5_2(reg_results)

    print("\n" + "=" * 60)
    print(f"All 8 figures saved to '{OUTPUT_DIR}/' directory.")
    print("=" * 60)


if __name__ == '__main__':
    main()
