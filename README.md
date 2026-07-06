# LCA_Allocation_Optimization_single.py Usage Guide

Scope of this document:
- How to prepare input and run the script
- How to read output files quickly
- Where to tune run-time parameters
- How to start from git clone

If you need optimization math/constraint-level detail, see:
- LCA_Allocation_Optimization_Model_Guide.md

## 1) Purpose
This script runs a single-cycle LCA allocation optimization.
It reads input data from an Excel workbook, solves an optimization model with OR-Tools (SCIP), and exports detailed Excel outputs.

Main script:
- LCA_Allocation_Optimization_single.py

## 2) Inputs (Excel sheets)
The script expects these sheets in the input workbook:
- df_re_final: product requirements by period
- df_idel_final: idle data by product/period
- df_in_final: initial inventory by product
- df_con_cost: conversion cost matrix (from product to product)
- df_cost: purchase/new buy cost by product

Current default input path in main:
- input/Demo_input_POR3.xlsx

## 3) Environment and dependencies
Required Python packages:
- pandas
- ortools
- openpyxl

## 4) Setup from git clone
1. Clone repository:
  - git clone https://github.com/SuphisaraLimsuthakul/Workshop_HSA_LCA_Allocation_Optimization.git
2. Move into project folder:
  - cd Workshop_HSA_LCA_Allocation_Optimization
3. Create virtual environment:
  - python -m venv .venv
4. Activate virtual environment (Windows PowerShell):
  - .venv\Scripts\Activate.ps1
5. Install dependencies:
  - pip install pandas ortools openpyxl
6. Move to working folder:
  - cd workshop

## 5) How to run
1. Open the script and confirm input/output paths in the main section.
2. Run the script.
  - python LCA_Allocation_Optimization_single.py
3. Review console summary.
4. Open generated Excel output for details.

## 6) Core model behavior (high level)
- Requirement for optimization model uses Value_Round.
- Requirement shown in output prefers Value (non-rounded) if available.
- Conversion is allowed only in periods with requirement.
- Buy has 2-period lead time.
- First two periods force shortage to 0.
- Objective minimizes total cost:
  - purchase cost + conversion cost

## 7) Generated outputs
The script writes timestamped files to workshop/output:
- LCA_Output_POR3_YYYYMMDD_HHMMSS.xlsx
- LCA_Model_POR3_YYYYMMDD_HHMMSS.lp

Important output sheets:
- Allocation_Summary
- Total_Cost_Summary
- LCA_Requirement_Summary
- Conversion_Details
- Conversion_Accounting_By_ToProduct
- Shortage_Details
- Product_* (one sheet per product)

## 8) Main functions (runtime view)
- build_data_from_demo_input(path): load and normalize input data
- solve_lca(data, allocation_timing_mode, lp_output_path): solve optimization model
- export_outputs(data, sol, output_path): write Excel outputs
- build_usage_and_cost_reports(data, sol): conversion/idle/shortage reports
- build_monthly_summary_tables(data, sol): summary tables for console/output

## 9) Key parameters to adjust
In the main block:
- path: input workbook path
- output_path: output Excel path pattern
- lp_output_path: LP export path pattern

In solve_lca:
- allocation_timing_mode:
  - requirement_month (default)
  - cost_optimal

## 10) Troubleshooting
If no optimal solution is found:
- Check requirement vs opening inventory in first two periods.
- Verify conversion cost and product mapping are valid.
- Inspect exported LP file for conflicting constraints.
- Ensure required sheets/columns exist and names are correct.

## 11) Typical workflow
1. Prepare/update Demo_input_POR3.xlsx.
2. Run script.
3. Check Total_Cost_Summary and Conversion_Details first.
4. Validate Allocation_Summary vs LCA_Requirement_Summary.
5. Iterate on input assumptions/costs and rerun.

## 12) Document boundary
This file intentionally avoids full mathematical notation and full constraint derivations.
Use LCA_Allocation_Optimization_Model_Guide.md for complete optimization details.
