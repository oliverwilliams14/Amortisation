library(readxl)
library(writexl)
library(ggplot2)
library(dplyr)
library(tidyr)
library(stringr)
library(purrr)
library(scales)
library(gridExtra)
library(tools)

# Calculate amortization rate based on future capex and LOM ounces
calculate_amortization_rate <- function(future_capex, lom_ounces) {
  if (lom_ounces == 0) {
    return(NA)  # Avoid division by zero
  }
  return(as.numeric(future_capex) / as.numeric(lom_ounces))
}

# Calculate expected expense by multiplying amortization rate by ounces mined
calculate_expected_expense <- function(amortization_rate, ounces_mined) {
  return(as.numeric(amortization_rate) * as.numeric(ounces_mined))
}

# Perform sensitivity analysis on amortization rate and expected expenses
sensitivity_analysis <- function(base_future_capex, base_lom_ounces, 
                                ounces_mined = NULL, variation = 0.20, steps = 5) {
  # Calculate variation percentages
  percentages <- seq(-variation, variation, length.out = 2 * variation / (variation / steps) + 1)
  
  # Create vectors for future capex and LOM ounces variations
  future_capex_variations <- base_future_capex * (1 + percentages)
  lom_ounces_variations <- base_lom_ounces * (1 + percentages)
  
  # Create percentage labels for the table
  percentage_labels <- paste0(as.integer(percentages * 100), "%")
  
  # Initialize matrices for results
  amort_matrix <- matrix(0, nrow = length(percentage_labels), ncol = length(percentage_labels))
  
  # Fill the matrix with amortization rates
  for (i in 1:length(percentage_labels)) {
    for (j in 1:length(percentage_labels)) {
      amort_matrix[i, j] <- calculate_amortization_rate(future_capex_variations[i], 
                                                        lom_ounces_variations[j])
    }
  }
  
  # Create data frames from matrices
  amort_results <- as.data.frame(amort_matrix)
  colnames(amort_results) <- percentage_labels
  rownames(amort_results) <- percentage_labels
  
  # Calculate expected expenses if ounces mined is provided
  expense_results <- NULL
  if (!is.null(ounces_mined)) {
    expense_matrix <- matrix(0, nrow = length(percentage_labels), ncol = length(percentage_labels))
    for (i in 1:length(percentage_labels)) {
      for (j in 1:length(percentage_labels)) {
        expense_matrix[i, j] <- calculate_expected_expense(amort_matrix[i, j], ounces_mined)
      }
    }
    
    expense_results <- as.data.frame(expense_matrix)
    colnames(expense_results) <- percentage_labels
    rownames(expense_results) <- percentage_labels
  }
  
  return(list(amort_results = amort_results, expense_results = expense_results))
}

# Read input data from an Excel file
read_input_excel <- function(file_path) {
  tryCatch({
    # Try to read the Excel file
    df <- readxl::read_excel(file_path)
    
    # Validate that required columns exist
    required_columns <- c('Project', 'Future_Capex', 'LOM_Ounces', 'Ounces_Mined')
    missing_columns <- required_columns[!required_columns %in% colnames(df)]
    
    if (length(missing_columns) > 0) {
      cat("Error: Missing required columns in the Excel file:", paste(missing_columns, collapse = ", "), "\n")
      cat("Please ensure your Excel file has columns named: Project, Future_Capex, LOM_Ounces, Ounces_Mined\n")
      return(NULL)
    }
    
    # Ensure numeric columns are numeric
    df$Future_Capex <- as.numeric(df$Future_Capex)
    df$LOM_Ounces <- as.numeric(df$LOM_Ounces)
    df$Ounces_Mined <- as.numeric(df$Ounces_Mined)
    
    # Drop rows with NA values in required numeric columns
    df <- df[complete.cases(df[, c('Future_Capex', 'LOM_Ounces', 'Ounces_Mined')]), ]
    
    if (nrow(df) == 0) {
      cat("Error: No valid data rows found after cleaning.\n")
      return(NULL)
    }
    
    return(df)
  }, error = function(e) {
    cat("Error reading Excel file:", e$message, "\n")
    return(NULL)
  })
}

# Create and save a heatmap visualization
create_heatmap <- function(data, title, filename, save_dir) {
  # Reshape the data for ggplot
  plot_data <- as.data.frame(data) %>%
    rownames_to_column(var = "Future_Capex") %>%
    pivot_longer(cols = -Future_Capex, 
                names_to = "LOM_Ounces", 
                values_to = "Value")
  
  # Create the heatmap
  p <- ggplot(plot_data, aes(x = LOM_Ounces, y = Future_Capex, fill = Value)) +
    geom_tile() +
    geom_text(aes(label = sprintf("%.2f", Value)), color = "black", size = 3) +
    scale_fill_gradient(low = "lightyellow", high = "steelblue") +
    labs(title = title,
         x = "LOM Ounces Variation",
         y = "Future Capex Variation") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5))
  
  # Save the plot
  filepath <- file.path(save_dir, filename)
  ggsave(filepath, p, width = 12, height = 10, dpi = 300)
  cat("Heatmap saved to:", filepath, "\n")
}

# Ask user for a file path and validate it exists
get_file_path <- function(prompt) {
  repeat {
    file_path <- trimws(readline(prompt))
    
    if (file_path == "") {
      return(NULL)
    }
    
    if (file.exists(file_path)) {
      return(file_path)
    } else {
      cat("Error: File '", file_path, "' does not exist.\n")
    }
  }
}

# Ask user for a directory path where files should be saved
get_save_path <- function() {
  repeat {
    save_dir <- trimws(readline("\nEnter the folder path to save results (or press Enter to use current directory): "))
    
    # Use current directory if input is empty
    if (save_dir == "") {
      return(getwd())
    }
    
    # Check if the directory exists
    if (dir.exists(save_dir)) {
      return(save_dir)
    } else {
      create_dir <- readline(paste0("Directory '", save_dir, "' doesn't exist. Create it? (y/n): "))
      if (tolower(create_dir) == 'y') {
        tryCatch({
          dir.create(save_dir, recursive = TRUE)
          cat("Created directory:", save_dir, "\n")
          return(save_dir)
        }, error = function(e) {
          cat("Error creating directory:", e$message, "\n")
        })
      } else {
        cat("Please enter a valid directory path.\n")
      }
    }
  }
}

# Main function
main <- function() {
  cat("Batch Sensitivity Analysis for Amortization Rate and Expected Expenses\n")
  cat("--------------------------------------------------------------------\n")
  
  # Get input Excel file path
  input_file <- get_file_path("Enter the path to your input Excel file: ")
  if (is.null(input_file)) {
    cat("Operation cancelled.\n")
    return(invisible(NULL))
  }
  
  # Read input data
  input_data <- read_input_excel(input_file)
  if (is.null(input_data)) {
    return(invisible(NULL))
  }
  
  # Get output directory
  output_dir <- get_save_path()
  if (is.null(output_dir)) {
    cat("Operation cancelled.\n")
    return(invisible(NULL))
  }
  
  # Process each row in the input data
  cat("\nProcessing", nrow(input_data), "projects from the input file...\n")
  
  for (i in 1:nrow(input_data)) {
    tryCatch({
      # Extract data from the row
      row <- input_data[i, ]
      project_name <- row$Project
      future_capex <- as.numeric(row$Future_Capex)
      lom_ounces <- as.numeric(row$LOM_Ounces)
      ounces_mined <- as.numeric(row$Ounces_Mined)
      
      cat("\nProcessing Project:", project_name, "\n")
      cat("  Future Capex: $", format(future_capex, big.mark = ",", scientific = FALSE), "\n", sep = "")
      cat("  LOM Ounces: ", format(lom_ounces, big.mark = ",", scientific = FALSE), "\n", sep = "")
      cat("  Ounces Mined: ", format(ounces_mined, big.mark = ",", scientific = FALSE), "\n", sep = "")
      
      # Create project-specific output directory
      project_dir <- file.path(output_dir, project_name)
      dir.create(project_dir, recursive = TRUE, showWarnings = FALSE)
      
      # Perform sensitivity analysis
      results <- sensitivity_analysis(future_capex, lom_ounces, ounces_mined)
      amort_results <- results$amort_results
      expense_results <- results$expense_results
      
      # Save results to Excel
      excel_filename <- paste0(project_name, "_sensitivity_analysis.xlsx")
      excel_path <- file.path(project_dir, excel_filename)
      
      # Create a list of data frames for the Excel file
      sheets <- list(
        "Amortization Rates" = amort_results,
        "Expected Expenses" = expense_results,
        "Input Summary" = data.frame(
          Parameter = c("Project", "Future Capex", "LOM Ounces", "Ounces Mined"),
          Value = c(project_name, future_capex, lom_ounces, ounces_mined)
        )
      )
      
      # Write to Excel
      writexl::write_xlsx(sheets, excel_path)
      cat("  Excel results saved to:", excel_path, "\n")
      
      # Create and save amortization rate heatmap
      create_heatmap(
        amort_results,
        paste0("Amortization Rate Sensitivity Analysis: ", project_name, " ($/ounce)"),
        paste0(project_name, "_amortization_sensitivity.png"),
        project_dir
      )
      
      # Create and save expected expense heatmap
      create_heatmap(
        expense_results,
        paste0("Expected Expense Sensitivity: ", project_name, " (", 
               format(ounces_mined, big.mark = ",", scientific = FALSE), " Ounces Mined)"),
        paste0(project_name, "_expense_sensitivity.png"),
        project_dir
      )
      
      cat("  Sensitivity analysis completed for", project_name, "\n")
      
    }, error = function(e) {
      cat("Error processing row", i, "(Project:", input_data$Project[i], "):", e$message, "\n")
    })
  }
  
  cat("\nBatch processing completed!\n")
}

# Run the main function
main()