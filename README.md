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
1. Clone this repo  
   ```bash
   git clone https://github.com/<you>/shiny-anova-app.git
   cd shiny-anova-app
