# SPDX-License-Identifier: GPL-2.0-or-later

# Load necessary libraries
library(shiny)
library(readxl)
library(tidyverse)
library(openxlsx)
library(emmeans)
library(multcomp)
library(DT)
library(car)

# Helper function to generate unique sheet names (for Excel)
generate_unique_sheet_name <- function(base_name, suffix, existing_sheets, max_sheet_name_length = 31) {
  available_space <- max_sheet_name_length - nchar(suffix)
  truncated_base <- substr(base_name, 1, available_space)
  sheet_name <- paste0(truncated_base, suffix)
  
  counter <- 1
  max_attempts <- 100
  while (sheet_name %in% existing_sheets && counter < max_attempts) {
    counter_suffix <- paste0("_", counter)
    truncated_base <- substr(base_name, 1, available_space - nchar(counter_suffix))
    sheet_name <- paste0(truncated_base, suffix, counter_suffix)
    counter <- counter + 1
  }
  if (counter == max_attempts) stop("Failed to generate a unique sheet name after 100 attempts.")
  sheet_name
}

# Helper function to sanitize variable names (wrap in backticks if needed)
sanitize_var <- function(var) {
  if (!grepl("^[a-zA-Z][a-zA-Z0-9_]*$", var)) {
    paste0("`", var, "`")
  } else {
    var
  }
}

# Define UI
ui <- fluidPage(
  titlePanel("ANOVA with Flexible EMMEANS"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Excel File", accept = ".xlsx"),
      uiOutput("sheet_selector"),
      radioButtons("response_mode", "Response Mode",
                   choices = c("Single Response" = "single",
                               "All Numeric Responses" = "all"),
                   selected = "single"),
      uiOutput("response_selector"),
      uiOutput("predictor_selector"),
      helpText("⚠️ Too many predictors may slow down the analysis. If performance suffers, consider reducing the number of predictors manually."),
      tags$hr(),
      h4("EMMEANS Specification"),
      helpText(
        "Select the main factor for estimated marginal means.",
        "Optionally, select a secondary factor (By Factor) to condition on."
      ),
      selectInput("emmeans_main", "Main Factor", choices = NULL),
      selectInput("emmeans_by", "By Factor (optional)", choices = c("None" = "None")),
      tags$hr(),
      selectInput("anova_method", "Select ANOVA Method",
                  choices = c("car::Anova Type I" = "type1",
                              "car::Anova Type II" = "type2",
                              "car::Anova Type III" = "type3"),
                  selected = "type1"),
      actionButton("run", "Run Analysis"),
      br(), br(),
      downloadButton("downloadData", "Download Results (Excel)"),
      br(), br(),
      tags$em("Note: Variable names will be sanitized if they aren’t valid R identifiers.")
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Data Preview", DTOutput("data_preview")),
        tabPanel("ANOVA Summary",
                 uiOutput("response_view_selector"),
                 verbatimTextOutput("results_table")),
        tabPanel("EMMEANS",
                 uiOutput("response_view_selector"),
                 DTOutput("emmeans_table")),
        tabPanel("Pairwise Comparisons",
                 uiOutput("response_view_selector"),
                 DTOutput("pairwise_table"))
      )
    )
  )
)

# Define server logic
server <- function(input, output, session) {
  
  ### Data Input & Setup ###
  data_file <- reactive({
    req(input$file)
    validate(
      need(grepl("\\.xlsx$", input$file$name, ignore.case = TRUE),
           "Please upload a valid .xlsx file.")
    )
    list(
      path = input$file$datapath,
      sheets = excel_sheets(input$file$datapath)
    )
  })
  
  output$sheet_selector <- renderUI({
    req(data_file())
    selectInput("sheet", "Select Sheet", choices = data_file()$sheets)
  })
  
  data_raw <- reactive({
    req(input$file, input$sheet)
    read_excel(data_file()$path, sheet = input$sheet) %>% as.data.frame()
  })
  
  output$response_selector <- renderUI({
    req(data_raw())
    numeric_vars <- names(data_raw())[sapply(data_raw(), is.numeric)]
    if (length(numeric_vars) == 0) return(tags$em("No numeric variables found in the sheet."))
    if (input$response_mode == "single") {
      selectInput("response", "Select Response Variable", choices = numeric_vars)
    } else {
      selectizeInput("response_all", "Select Numeric Responses to Analyze",
                     choices = numeric_vars, selected = numeric_vars, multiple = TRUE)
    }
  })
  
  output$predictor_selector <- renderUI({
    req(data_raw())
    preds <- if (input$response_mode == "single") {
      setdiff(names(data_raw()), input$response)
    } else {
      setdiff(names(data_raw()), input$response_all)
    }
    if (length(preds) == 0) return(tags$em("No predictor variables available."))
    selectizeInput("predictors", "Select Predictor Variables",
                   choices = preds, selected = preds, multiple = TRUE)
  })
  
  observe({
    req(input$predictors)
    updateSelectInput(session, "emmeans_main", choices = input$predictors,
                      selected = input$predictors[1])
    updateSelectInput(session, "emmeans_by", choices = c("None" = "None", input$predictors),
                      selected = "None")
  })
  
  output$data_preview <- renderDT({
    req(data_raw())
    datatable(data_raw(), options = list(pageLength = 10))
  })
  
  ### Analysis ###
  analysis_results <- eventReactive(input$run, {
    req(data_raw(), input$predictors)
    df <- data_raw()
    df[input$predictors] <- lapply(df[input$predictors], as.factor)
    
    responses <- if (input$response_mode == "single") c(input$response) else input$response_all
    
    main_spec <- input$emmeans_main
    by_spec <- input$emmeans_by
    em_formula <- if (by_spec != "None") {
      as.formula(paste("~", main_spec, "|", by_spec))
    } else {
      as.formula(paste("~", main_spec))
    }
    
    results <- vector("list", length(responses))
    names(results) <- responses
    
    withProgress(message = "Running analyses...", value = 0, {
      total <- length(responses)
      for (i in seq_along(responses)) {
        resp <- responses[i]
        incProgress(1/total, detail = paste("Processing response", resp))
        
        s_response <- sanitize_var(resp)
        s_preds <- sapply(input$predictors, sanitize_var)
        model_formula <- as.formula(paste(s_response, "~", paste(s_preds, collapse = " * ")))
        model <- aov(model_formula, data = df)
        
        anova_obj <- switch(input$anova_method,
                            type1 = Anova(model, type = "I"),
                            type2 = Anova(model, type = "II"),
                            type3 = Anova(model, type = "III"))
        anova_text <- capture.output(anova_obj)
        
        emm <- tryCatch(emmeans(model, specs = em_formula), error = function(e) NULL)
        em_res <- if (!is.null(emm)) {
          cld_out <- cld(emm, alpha = 0.05, Letters = letters, details = TRUE)
          as.data.frame(cld_out$emmeans)
        } else data.frame(Message = "Error in emmeans")
        
        pw_res <- if (!is.null(emm)) {
          as.data.frame(pairs(emm))
        } else data.frame(Message = "Error in pairwise comparisons")
        
        results[[resp]] <- list(
          anova_text = anova_text,
          emmeans = em_res,
          pairwise = pw_res
        )
      }
    })
    list(single = (length(responses) == 1), results = results)
  })
  
  output$response_view_selector <- renderUI({
    res <- analysis_results()
    if (!res$single) selectInput("selected_response", "Select Response to View",
                                 choices = names(res$results))
  })
  
  output$results_table <- renderPrint({
    res <- analysis_results()
    if (res$single) cat(res$results[[1]]$anova_text, sep = "
") else {
  req(input$selected_response)
  cat(res$results[[input$selected_response]]$anova_text, sep = "
")
}
  })
  
  output$emmeans_table <- renderDT({
    res <- analysis_results()
    tbl <- if (res$single) res$results[[1]]$emmeans else res$results[[input$selected_response]]$emmeans
    datatable(tbl, options = list(pageLength = 10))
  })
  
  output$pairwise_table <- renderDT({
    res <- analysis_results()
    tbl <- if (res$single) res$results[[1]]$pairwise else res$results[[input$selected_response]]$pairwise
    datatable(tbl, options = list(pageLength = 10))
  })
  
  ### Download Handler ###
  output$downloadData <- downloadHandler(
    filename = function() {
      pref <- if (analysis_results()$single) "ANOVA_Results_" else "ANOVA_Results_All_"
      paste0(pref, input$sheet, ".xlsx")
    },
    content = function(file) {
      wb <- createWorkbook()
      existing_sheets <- c()
      res <- analysis_results()
      for (resp in names(res$results)) {
        out <- res$results[[resp]]
        sheet1 <- generate_unique_sheet_name(resp, "_ANOVA", existing_sheets)
        existing_sheets <- c(existing_sheets, sheet1)
        addWorksheet(wb, sheet1)
        writeData(wb, sheet1, data.frame(Line = out$anova_text), colNames = FALSE)
        sheet2 <- generate_unique_sheet_name(resp, "_EMMEANS", existing_sheets)
        existing_sheets <- c(existing_sheets, sheet2)
        addWorksheet(wb, sheet2)
        writeData(wb, sheet2, out$emmeans)
        sheet3 <- generate_unique_sheet_name(resp, "_Pairwise", existing_sheets)
        existing_sheets <- c(existing_sheets, sheet3)
        addWorksheet(wb, sheet3)
        writeData(wb, sheet3, out$pairwise)
      }
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# Run the application
shinyApp(ui, server)
