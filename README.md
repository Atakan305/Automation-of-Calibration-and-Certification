# Automation of Calibration and Certification Generic Tasks
Generic Python-based automation framework for calibration and certification workflows. It extracts calibration results and metadata from Excel certificates, builds historical traceability datasets, performs statistical/error analysis, generates visual reports, and accelerates large-scale archive processing while reducing manual work and human error. 

# What the Script Does

The program scans calibration certificate archives and:

- finds certificates belonging to the same instrument or equipment chain
- extracts calibration tables automatically
- reads reference, reading, error, and uncertainty values
- reconstructs calibration history over time
- generates comparison datasets
- creates trend plots and summary reports

The workflow is especially useful when the same instrument is calibrated multiple times across different years and you want to quickly evaluate drift, repeatability, or long-term behavior.

---

# Main Capabilities

## 1. Fast Archive Search

The project creates a cached index of the archive instead of rescanning every folder during every execution.

This makes a major difference when working with:

- large shared company folders
- network drives
- multi-year certificate archives
- thousands of Excel certificates

---

## 2. Flexible Certificate Parsing

Real company (mostly laboratory-based) certificates are sometimes not standardized perfectly.

Some use different:

- sheet names
- metadata labels
- table structures
- languages
- layouts

Because of that, the parser was intentionally written using heuristic-based detection instead of fixed cell positions.

The script tries to identify:

- result tables
- metadata regions
- date fields
- serial numbers
- uncertainty columns
- calibration points

without depending on one exact template.

---

## 3. Historical Calibration Reconstruction

The script groups certificates belonging to the same instrument by using identifiers extracted from:

- file names
- workbook contents
- serial number fields
- equipment labels

It then rebuilds the complete calibration history automatically.

Example:

```text
2022 → Calibration #1
2024 → Calibration #2
2026 → Calibration #3
```

This makes long-term comparison extremely fast.

---

## 4. Nominal Value Normalization

Calibration references are often stored as values like:

```text
24.99
40.01
80.03
```

but operationally they correspond to nominal points:

```text
25 °C
40 °C
80 °C
```

The script automatically normalizes these values using configurable tolerances so historical comparisons remain consistent.

---

## Generated Outputs

The project exports:

- raw calibration history
- metadata history
- pivot tables
- summary statistics
- Excel reports
- calibration trend plots

The generated outputs are intended to be both machine-readable and easy to inspect manually.

---

## Typical Use Cases

This project is useful for:

- certification traceability
- ISO/IEC 17025 workflows
- calibration-like laboratories
- metrology environments
- internal quality systems
- instrument drift monitoring
- historical performance analysis

---

## Libraries

Main libraries used:

```text
pandas
numpy
matplotlib
openpyxl
xlrd
```

---

## Installation

```bash
pip install pandas numpy matplotlib openpyxl xlrd==2.0.1
```

---

## Running the Script

```bash
python main.py
```

The program will ask for:

- archive location
- identifiers / serial numbers
- output location
- analysis settings

---

## Example Output Structure

```text
output/
│
├── raw_calibration_history.csv
├── certificate_metadata.csv
├── calibration_automation_report.xlsx
├── pivot_error.csv
├── summary_by_nominal.csv
└── plots/
```

---

## Notes

The project was built around real workflows rather than synthetic example datasets.

For that reason, the code focuses heavily on the main aspects as follows:

- robustness
- imperfect Excel structures
- flexible parsing
- traceability
- practical archive handling

instead of relying on perfectly formatted templates.

Some certificate formats may still require small adjustments depending on their structure.

---

## License

Open for educational, research, and workflow automation purposes. 
