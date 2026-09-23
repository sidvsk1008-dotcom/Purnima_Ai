# Purnima AI

Purnima AI is a Python-based automation project for transforming standardized business workbook data into structured, validated and auditable accounting outputs.

## Purpose

The project follows a controlled data-processing pipeline:

Standardized Excel
↓
Input Safety Gate
↓
Extraction
↓
Validation
↓
Classification
↓
Business Rules
↓
Calculations
↓
Reconciliation
↓
Audit / Diagnostics
↓
Standardized Output

## Engineering Principles

- Accept exactly one explicit source workbook for a processing run.
- Never scan an input directory and silently choose a workbook.
- Preserve source information.
- Never silently invent missing information.
- Separate extraction from classification and calculations.
- Keep business rules explicit and testable.
- Require human confirmation when configured validation rules identify ambiguity.
- Maintain diagnostic information for traceability.
- Keep milk supplier metadata separate from business-unit destination logic.
- Use regression tests and verified reference data to detect unintended changes.

## Technology

- Python
- OpenPyXL
- JSON
- Excel
- Git
- GitHub

## Processing Layers

### 1. Input Safety Gate

Verifies that exactly one valid Excel workbook has been supplied.

### 2. Extraction

Reads the standardized workbook and converts its rows into a canonical internal representation.

### 3. Validation

Checks structural and business-rule requirements before processing continues.

### 4. Classification

Determines the business meaning of each record, such as purchase, expense, sale or payment/control record.

### 5. Business Rules

Applies explicit Purnima accounting rules without mixing them into raw Excel extraction.

### 6. Calculations

Calculates accounting totals from classified data.

### 7. Reconciliation

Checks related values and identifies inconsistencies.

### 8. Audit and Diagnostics

Records how data was interpreted and why records were included or excluded from calculations.

## Engineering Workflow

SORT
→ SEGREGATE
→ MODEL
→ BOUND
→ CONTRACT
→ MODULARIZE
→ CODE
→ TEST
→ DEPLOY / OPERATE

## Repository Structure

```text
Purnima_Ai/
├── main.py
├── rules.json
├── input/
├── output/
├── archive/
├── exceptions/
└── tests/
