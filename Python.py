import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import glob
from datetime import datetime

def calculate_amortization_rate(future_capex, lom_ounces):
    """
    Calculate amortization rate based on future capex and LOM ounces.
    
    Args:
        future_capex (float): Future capital expenditure in dollars
        lom_ounces (float): Life of Mine ounces
        
    Returns:
        float: Amortization rate ($/ounce)
    """
    if lom_ounces == 0:
        return np.nan  # Avoid division by zero
    return float(future_capex) / float(lom_ounces)

def calculate_expected_expense(amortization_rate, ounces_mined):
    """
    Calculate expected expense by multiplying amortization rate by ounces mined.
    
    Args:
        amortization_rate (float): Amortization rate in $/ounce
        ounces_mined (float): Number of ounces mined
        
    Returns:
        float: Expected expense in dollars
    """
    return float(amortization_rate) * float(ounces_mined)

def sensitivity_analysis(base_future_capex, base_lom_ounces, ounces_mined=None, variation=0.20, steps=5):
    """
    Perform sensitivity analysis on amortization rate and expected expenses.
    
    Args:
        base_future_capex (float): Base value for future capex
        base_lom_ounces (float): Base value for LOM ounces
        ounces_mined (float, optional): Number of ounces mined for expense calculation
        variation (float): Variation percentage (default 0.20 for ±20%)
        steps (int): Number of steps between variations (default 5 for 5% intervals)
        
    Returns:
        tuple: (amortization_rates DataFrame, expected_expenses DataFrame or None)
    """
    # Calculate variation percentages
    percentages = np.linspace(-variation, variation, int(2 * variation / (variation / steps)) + 1)
    
    # Create arrays for future capex and LOM ounces variations
    future_capex_variations = [base_future_capex * (1 + p) for p in percentages]
    lom_ounces_variations = [base_lom_ounces * (1 + p) for p in percentages]
    
    # Create percentage labels for the table
    percentage_labels = [f"{int(p * 100)}%" for p in percentages]
    
    # Create numeric matrices for the heatmaps
    amort_matrix = np.zeros((len(percentage_labels), len(percentage_labels)))
    
    # Fill the matrix with amortization rates
    for i, fc in enumerate(future_capex_variations):
        for j, lo in enumerate(lom_ounces_variations):
            amort_matrix[i, j] = calculate_amortization_rate(fc, lo)
    
    # Create the DataFrame from the numeric matrix
    amort_results = pd.DataFrame(amort_matrix, index=percentage_labels, columns=percentage_labels)
    
    # Label the index and columns
    amort_results.index.name = "Future Capex Variation"
    amort_results.columns.name = "LOM Ounces Variation"
    
    # Calculate expected expenses if ounces mined is provided
    expense_results = None
    if ounces_mined is not None:
        expense_matrix = np.zeros((len(percentage_labels), len(percentage_labels)))
        for i in range(len(percentage_labels)):
            for j in range(len(percentage_labels)):
                expense_matrix[i, j] = calculate_expected_expense(amort_matrix[i, j], ounces_mined)
        
        expense_results = pd.DataFrame(expense_matrix, index=percentage_labels, columns=percentage_labels)
        expense_results.index.name = "Future Capex Variation"
        expense_results.columns.name = "LOM Ounces Variation"
    
    return amort_results, expense_results

def read_input_excel(file_path):
    """
    Read input data from an Excel file.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        pd.DataFrame: DataFrame containing the input data
    """
    try:
        # Try to read the Excel file
        df = pd.read_excel(file_path)
        
        # Validate that required columns exist
        required_columns = ['Project', 'Future_Capex', 'LOM_Ounces', 'Ounces_Mined']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Error: Missing required columns in the Excel file: {', '.join(missing_columns)}")
            print("Please ensure your Excel file has columns named: Project, Future_Capex, LOM_Ounces, Ounces_Mined")
            return None
        
        # Ensure numeric columns are numeric
        for col in ['Future_Capex', 'LOM_Ounces', 'Ounces_Mined']:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Drop rows with NaN values in required numeric columns
        df = df.dropna(subset=['Future_Capex', 'LOM_Ounces', 'Ounces_Mined'])
        
        if df.empty:
            print("Error: No valid data rows found after cleaning.")
            return None
            
        return df
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def create_heatmap(data, title, filename, save_dir):
    """
    Create and save a heatmap visualization.
    
    Args:
        data (pd.DataFrame): Data to visualize
        title (str): Title for the heatmap
        filename (str): Filename to save the heatmap
        save_dir (str): Directory to save the file
    """
    plt.figure(figsize=(12, 10))
    sns.heatmap(data, annot=True, cmap="YlGnBu", fmt=".2f")
    plt.title(title)
    plt.tight_layout()
    
    # Save the figure
    filepath = os.path.join(save_dir, filename)
    plt.savefig(filepath)
    print(f"Heatmap saved to: {filepath}")
    plt.close()  # Close the figure to free memory

def get_file_path(prompt):
    """
    Ask user for a file path and validate it exists.
    
    Args:
        prompt (str): Prompt message for the user
        
    Returns:
        str: Validated file path or None if cancelled
    """
    while True:
        file_path = input(prompt).strip()
        
        if not file_path:
            return None
            
        if os.path.exists(file_path):
            return file_path
        else:
            print(f"Error: File '{file_path}' does not exist.")

def get_save_path():
    """
    Ask user for a directory path where files should be saved.
    Validates the path and returns it, or returns current directory if empty input.
    """
    while True:
        save_dir = input("\nEnter the folder path to save results (or press Enter to use current directory): ").strip()
        
        # Use current directory if input is empty
        if not save_dir:
            return os.getcwd()
        
        # Check if the directory exists
        if os.path.isdir(save_dir):
            return save_dir
        else:
            create_dir = input(f"Directory '{save_dir}' doesn't exist. Create it? (y/n): ")
            if create_dir.lower() == 'y':
                try:
                    os.makedirs(save_dir, exist_ok=True)
                    print(f"Created directory: {save_dir}")
                    return save_dir
                except Exception as e:
                    print(f"Error creating directory: {e}")
            else:
                print("Please enter a valid directory path.")

def main():
    print("Batch Sensitivity Analysis for Amortization Rate and Expected Expenses")
    print("--------------------------------------------------------------------")
    
    # Get input Excel file path
    input_file = get_file_path("Enter the path to your input Excel file: ")
    if not input_file:
        print("Operation cancelled.")
        return
    
    # Read input data
    input_data = read_input_excel(input_file)
    if input_data is None:
        return
    
    # Get output directory
    output_dir = get_save_path()
    if not output_dir:
        print("Operation cancelled.")
        return
    
    # Process each row in the input data
    print(f"\nProcessing {len(input_data)} projects from the input file...")
    
    for index, row in input_data.iterrows():
        try:
            # Extract data from the row
            project_name = row['Project']
            future_capex = float(row['Future_Capex'])
            lom_ounces = float(row['LOM_Ounces'])
            ounces_mined = float(row['Ounces_Mined'])
            
            print(f"\nProcessing Project: {project_name}")
            print(f"  Future Capex: ${future_capex:,.2f}")
            print(f"  LOM Ounces: {lom_ounces:,.2f}")
            print(f"  Ounces Mined: {ounces_mined:,.2f}")
            
            # Create project-specific output directory
            project_dir = os.path.join(output_dir, f"{project_name}")
            os.makedirs(project_dir, exist_ok=True)
            
            # Perform sensitivity analysis
            amort_results, expense_results = sensitivity_analysis(
                future_capex, lom_ounces, ounces_mined
            )
            
            # Convert to float to ensure proper display
            amort_results = amort_results.astype(float)
            expense_results = expense_results.astype(float)
            
            # Save results to Excel
            excel_filename = f"{project_name}_sensitivity_analysis.xlsx"
            excel_path = os.path.join(project_dir, excel_filename)
            
            with pd.ExcelWriter(excel_path) as writer:
                amort_results.to_excel(writer, sheet_name="Amortization Rates")
                expense_results.to_excel(writer, sheet_name="Expected Expenses")
                
                # Add an input summary sheet
                summary_df = pd.DataFrame({
                    'Parameter': ['Project', 'Future Capex', 'LOM Ounces', 'Ounces Mined'],
                    'Value': [project_name, future_capex, lom_ounces, ounces_mined]
                })
                summary_df.to_excel(writer, sheet_name="Input Summary", index=False)
            
            print(f"  Excel results saved to: {excel_path}")
            
            # Create and save amortization rate heatmap
            create_heatmap(
                amort_results,
                f"Amortization Rate Sensitivity Analysis: {project_name} ($/ounce)",
                f"{project_name}_amortization_sensitivity.png",
                project_dir
            )
            
            # Create and save expected expense heatmap
            create_heatmap(
                expense_results,
                f"Expected Expense Sensitivity: {project_name} ({ounces_mined:,.0f} Ounces Mined)",
                f"{project_name}_expense_sensitivity.png",
                project_dir
            )
            
            print(f"  Sensitivity analysis completed for {project_name}")
            
        except Exception as e:
            print(f"Error processing row {index+1} (Project: {row.get('Project', 'Unknown')}): {e}")
    
    print("\nBatch processing completed!")

if __name__ == "__main__":
    main()
