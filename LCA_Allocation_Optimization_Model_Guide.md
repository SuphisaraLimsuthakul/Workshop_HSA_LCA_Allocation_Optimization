# Optimization Model Guide: Conversion Planning

This document explains the `solve_lca` model in `LCA_Allocation_Optimization_single.py` and maps each mathematical rule to practical business conversion behavior.

Scope of this document:
- Optimization-focused deep dive (variables, constraints, objective)
- Business meaning of each optimization rule

For run instructions and day-to-day usage, see:
- workshop/LCA_Allocation_Optimization_single_USAGE.md

It covers:
- Decision Variables
- Constraint System
- Objective Function
- Business interpretation of conversion policies

It does not focus on:
- Python environment setup details
- Basic execution walkthrough

---

## 1) Business Purpose

The model finds the lowest-cost monthly plan under operational conversion rules:
1. Use existing inventory efficiently
2. Convert only on allowed source-destination pairs
3. Allow conversion inflow only in months with demand
4. Enforce routing discipline (destination at most one source per month + source cooldown)
5. Purchase follows a 2-month lead time before the shortage month

---

## 2) Sets, Indices, and Parameters

### 2.1 Sets / Indices
- Products: all products
- Months: planning periods (1,2,3,...)

### 2.2 Key Parameters
- $R_{p,t}$: requirement of product $p$ in month $t$
- $A0_p$: opening inventory before horizon
- $IdleCapProd_{p,t}$: idle capacity / idle stock baseline
- $ConvCost_{q,p}$: unit conversion cost from source $q$ to destination $p$
- $NewBuyCost_p$: unit purchase cost of product $p$
- $AllowedConvPairs$: allowed conversion set $(q,p)$
- $M$: Big-M for linking route activation and flow

Data notes:
- idle-to-idle conversion pairs are excluded in preprocessing
- if purchase cost is missing, the current code defaults `NewBuyCost` to 0
- opening inventory used by optimization is integerized from raw opening inventory (reports can preserve raw decimals in month 1 display)
- month-1 idle stock is added into opening inventory before optimization
- month-1 requirement is capped to opening inventory when requirement exceeds available opening stock
- cost-family mapping is applied in preprocessing so product variants under the same family (e.g., BlueWhale*, Dorado*) can share mapped cost rows

---

## 3) Decision Variables

All are nonnegative integer variables, except $z$ (binary):

1. $I_{p,t}$: inventory of product $p$ in month $t$
2. $A_{p,t}$: allocatable/available quantity of product $p$ in month $t$
3. $Sur_{p,t}$: surplus in month $t$
4. $Short_{p,t}$: shortage in month $t$
5. $Buy_{p,t}$: purchased quantity in month $t$
6. $x_{p,q,t}$: conversion quantity received as destination $p$ from source $q$ in month $t$
7. $z_{p,q,t}$: route activation binary for conversion $q \to p$ in month $t$
8. $IdleInv_{p,t}$: idle inventory ledger value
9. $y_{p,t}$: idle usage quantity in month $t$

Index note:
- In code, `x[(p,q,t)]` means flow from source $q$ to destination $p$ in month $t$.

---

## 4) Constraint System

## (1) Destination conversion gating by monthly requirement

For each destination product $p$ and month $t$:
$$
\sum_{q:(q,p)\in Allowed} x_{p,q,t}=0 \quad \text{if } R_{p,t}=0
$$
$$
\sum_{q:(q,p)\in Allowed} x_{p,q,t} \le R_{p,t} \quad \text{if } R_{p,t}>0
$$

Business meaning:
- No conversion inflow into months without demand
- Prevents early pre-conversion into zero-demand periods

---

## (2) Monthly inventory balance

For each product $p$:

First month:
$$
I_{p,1}=A0_p+\text{Inflow}_{p,1}-\text{Outflow}_{p,1}
$$

Second month:
$$
I_{p,2}=I_{p,1}+\text{Inflow}_{p,2}-\text{Outflow}_{p,2}
$$

Month 3 onward:
$$
I_{p,t}=I_{p,t-1}+\text{Inflow}_{p,t}-\text{Outflow}_{p,t}+Buy_{p,t-2},\quad t\ge 3
$$

Where:
$$
\text{Inflow}_{p,t}=\sum_{q:(q,p)\in Allowed} x_{p,q,t}
$$
$$
\text{Outflow}_{p,t}=\sum_{q:(p,q)\in Allowed} x_{q,p,t+1}
$$

Business meaning:
- Inflow arrives in month $t$
- Source inventory is deducted one month before conversion completion (for flows completing at $t+1$)
- Purchases are received with a 2-month lead time ($Buy_{p,t-2}$ arrives in month $t$)

---

## (3) Allocation definition
$$
A_{p,t}=I_{p,t}
$$

---

## (4) Surplus / shortage balance
$$
Sur_{p,t}-Short_{p,t}=A_{p,t}-R_{p,t}
$$

---

## (4.1) Two-month advance buy linkage

For order months that can map to an in-horizon shortage month:
$$
Buy_{p,t}=Short_{p,t+2}
$$

For the last two order months:
$$
Buy_{p,t}=0
$$

Without pre-horizon buy orders:
$$
Short_{p,1}=0,\quad Short_{p,2}=0
$$

Business meaning:
- Shortage in month $t$ must be covered by buy planned at month $t-2$
- The first two months cannot be short because there is no pre-horizon buy decision variable

---

## (4.2) Shortage upper bound by requirement
$$
Short_{p,t}\le R_{p,t}
$$

Business meaning:
- If requirement is zero, shortage must be zero
- Prevents artificial Sur/Short inflation when buy cost is zero

---

## (5) Idle usage cap
$$
y_{p,t}\le IdleInv_{p,t}
$$

---

## (5.1) Idle inventory ledger balance

For non-idle-ledger products:
$$
IdleInv_{p,t}=0
$$

For idle-ledger-tracked products:
- First month
$$
IdleInv_{p,1}=IdleCapProd_{p,1}+\text{ConvIn}_{p,1}-\text{ConvOut}_{p,1}
$$
- Later months
$$
IdleInv_{p,t}=IdleInv_{p,t-1}+\text{ConvIn}_{p,t}-\text{ConvOut}_{p,t}
$$

with the same source deduction timing as the main inventory ledger.

---

## (6) Source conversion cap by inventory from two months ago

- First two months:
$$
\sum_p x_{p,q,t}=0,\quad t\in\{1,2\}
$$
- Month 3 onward:
$$
\sum_p x_{p,q,t}\le I_{q,t-2}
$$

Business meaning:
- Source-side conversion feasibility is linked to historical inventory ($t-2$)

---

## (7) Idle usage cap by two-month lag

- First two months:
$$
y_{p,t}=0
$$
- Month 3 onward:
$$
y_{p,t}\le IdleCapProd_{p,t-2}
$$

---

## (8) Route activation and destination-side single-source rule

Route/flow linking:
$$
x_{p,q,t}\le M\cdot z_{p,q,t},\quad x_{p,q,t}\ge z_{p,q,t}
$$

Destination receives from at most one source in the same month:
$$
\sum_{q\in Src(p)} z_{p,q,t}\le 1
$$

Business meaning:
- A route binary is turned on only when quantity is shipped
- One destination can be fed by only one source per month
- A source may still feed multiple destinations in the same month (subject to inventory and cooldown constraints)

---

## (8.1) Optional anti-prebuild rule (requirement-month mode)

When `allocation_timing_mode = requirement_month`:
$$
Sur_{p,t}\le M\cdot(1-conv\_to\_dest\_active_{p,t})
$$

Business meaning:
- If destination $p$ receives conversion in month $t$, that month cannot end with surplus.
- This limits conversion to requirement-timing behavior rather than early pre-build.

---

## (8.2) Cooldown (no consecutive-month conversion by same source)
$$
\sum_{p\in Dest(q)} z_{p,q,t}+\sum_{p\in Dest(q)} z_{p,q,t+1}\le 1
$$

---

## 5) Objective Function

The model minimizes total cost:
$$
\min \left[
\sum_{p,t} NewBuyCost_p\cdot Buy_{p,t}
+ \sum_{(q,p)\in Allowed,\ t} ConvCost_{q,p}\cdot x_{p,q,t}
\right]
$$

Business meaning:
- Optimally trade off purchase versus conversion under all feasibility and policy constraints

---

## 6) Business Interpretation Highlights

1. Conversion can enter destination only when that month has demand
2. Source deduction appears one month before conversion completion
3. Buy/shortage follows a 2-month advance-planning policy
4. Source conversion capacity is still controlled by $t-2$ inventory policy
5. Early-month rules block conversion and idle usage until lag conditions are valid
6. Operational discipline is enforced by destination-side single-source rule plus source cooldown logic
7. Shortage is bounded by requirement, avoiding fake large purchases
8. In `requirement_month` mode, conversion into a month is blocked from creating surplus in that same month

---

## 7) Code-to-Business Mapping (for LCA_Allocation_Optimization_single.py)

- `x[(p,q,t)]`: conversion quantity from source q to destination p at month t
- `z[(p,q,t)]`: route activation switch for q->p
- `I[(p,t)]`: available inventory in month t
- `A[(p,t)]`: allocation in month t
- `Short[(p,t)]`: shortage amount
- `Buy[(p,t)]`: purchase amount
- `IdleInv[(p,t)]`: idle ledger balance
- `dests_by_source[q]`: destination list for source q
- `srcs_by_dest[p]`: source list that can convert into p

---

## 8) Executive Summary

This is a cost-optimal conversion planning model with destination-demand gating, 2-month purchase lead-time control, source feasibility controls, destination-side single-source routing with source cooldown, optional requirement-month anti-prebuild behavior, and anti-degeneracy shortage bounds, making results both financially optimal and operationally explainable.