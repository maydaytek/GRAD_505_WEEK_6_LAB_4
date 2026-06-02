# GRAD 505 - Week 6 Lab 4: Risk Assessment in Food Science

Cumulative pesticide exposure risk analysis using consumption survey data and pesticide residue measurements from Thai markets (Wanwimolruk et al. 2015, 2017).

## Overview

The analysis merges three datasets to quantify monthly pesticide exposure per subject and compare risk across eight demographic groups (sex x age class):

| Dataset | File | Description |
|---|---|---|
| Consumption | `MATERIALS/Consumption1.xlsx` | Food intake (g/occasion) and frequency (per month) for 7,267 subjects |
| Pesticide residues | `MATERIALS/Pesticide Residue2.xlsx` | Pesticide concentrations (ug/kg) in Chinese kale (2015) and cabbage (2017) |
| Toxicity | `MATERIALS/LD50.xlsx` | LD50 values (mg/kg) for 5 organophosphate pesticides used as relative potency factors |

**Target pesticides:**
- Chinese kale: Diazinon, Dimethoate, Malathion, Profenofos
- Cabbage: Diazinon, Dichlorvos, Dimethoate

## Methodology

Risk is calculated as:

```
Risk = sum_k( exposure_k x toxicity_k )
```

Where:
- **Amount_Eaten** (g/month) = consumption x frequency
- **TER** (Toxicity Exposure Ratio) = mean pesticide residue (mg/kg) x LD50 (mg/kg)
- **EXP** = TER x Amount_Eaten / body_weight
- **SUM_EXP** = sum of EXP across all pesticides for each vegetable
- **Relative exposure (%)** = (SUM_EXP_Total / 30.44 days) / reference dose x 100

The reference dose used for normalization is **0.0025 mg/kg/day** (Disulfoton, U.S. EPA).

## Usage

Requires Python 3 with `pandas`, `numpy`, and `openpyxl`:

```bash
pip install pandas numpy openpyxl
python lab4.py
```

Outputs are written to `outputs/`.

## Outputs

| File | Description |
|---|---|
| `outputs/Pesticide_Residue_Updated.xlsx` | Original residue sheets with mean rows appended (ug/kg and mg/kg) |
| `outputs/Pesticide_Residue_Subset.xlsx` | Trimmed to target pesticides only, with mean row |
| `outputs/Consumption_Revised.xlsx` | Full merged dataset: all subjects with Amount_Eaten, MC, RPF, TER, EXP, SUM_EXP, and relative exposure columns |
| `outputs/Risk_Summary.csv` | Mean, SD, min, max relative exposure (%) ranked by demographic group |

## Results Summary

Overall mean relative exposure across all subjects: **~16.8%** of the Disulfoton reference dose (SD ~26.9%).

Group rankings (highest to lowest mean relative exposure):

| Rank | Group | Mean % of Ref Dose |
|---|---|---|
| 1 | Female < 10 | 20.5% |
| 2 | Male < 10 | 19.1% |
| 3 | Female 18-39 | 18.9% |
| 4 | Female 10-17 | 18.9% |
| 5 | Male 18-39 | 18.9% |
| 6 | Male 10-17 | 17.3% |
| 7 | Female >= 40 | 12.1% |
| 8 | Male >= 40 | 12.1% |

Children under 10 show the highest mean risk, consistent with the hypothesis that younger age groups face greater exposure relative to body weight. Adults 40 and older show the lowest risk across both sexes.

## References

- Wanwimolruk, S. et al. (2015). *Food Science & Nutrition.*
- Wanwimolruk, S. et al. (2017). *Food Science & Nutrition.*
- U.S. EPA Reference Dose for Disulfoton: 0.0025 mg/kg/day
