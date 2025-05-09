SPDX-License-Identifier: GPL-2.0-or-later

# ANOVA with Flexible EMMEANS

This Shiny app performs ANOVA (Type I/II/III), EMMEANS and COMPACT LETTER DISPLAY analyses on uploaded Excel data.

## 🎯 Features
- Upload `.xlsx` files and select sheets  
- Choose single or multiple numeric responses  
- Specify predictors and EMMEANS factors  
- Download results as Excel

## ⚙️ Prerequisites
- R ≥ 3.0.2  
- Packages: `shiny` (GPL-3), `readxl` (MIT), `tidyverse` (MIT), `openxlsx` (MIT),  
  `emmeans` (GPL-2|GPL-3), `multcomp` (GPL-2), `DT` (GPL-3), `car` (GPL-2|GPL-3)

## 🚀 Usage
1. **Clone the repo**  
   ```bash
   git clone https://github.com/<you>/shiny-anova-app.git
   cd shiny-anova-app

2. **Start the app in R**
    ```bash
    shiny::runApp("app.R")

3. **Prepare your input Excel file**
Your spreadsheet must have three distinct sections, in this order:

| Section             | Description                                                                               |
| ------------------- | ----------------------------------------------------------------------------------------- |
| **Row labels**      | First column: unique IDs (e.g. gene or sample names)                                      |
| **Row annotations** | One or more columns immediately after: categorical factors (e.g. “genotype”, “condition”) |
| **Numeric data**    | All remaining columns: measurement values; headers are variable names (e.g. metabolites)  |
