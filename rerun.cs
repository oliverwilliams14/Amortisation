using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using ScottPlot;
using ScottPlot.Palettes;
using System.Drawing;

// For EPPlus licensing
[assembly: OfficeOpenXml.ExcelPackage.LicenseContext(OfficeOpenXml.LicenseContext.NonCommercial)]

namespace SensitivityAnalysis
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Batch Sensitivity Analysis for Amortization Rate and Expected Expenses");
            Console.WriteLine("--------------------------------------------------------------------");

            // Get input Excel file path
            string inputFile = GetFilePath("Enter the path to your input Excel file: ");
            if (string.IsNullOrEmpty(inputFile))
            {
                Console.WriteLine("Operation cancelled.");
                return;
            }

            // Read input data
            List<ProjectData> projects = ReadInputExcel(inputFile);
            if (projects == null || projects.Count == 0)
            {
                return;
            }

            // Get output directory
            string outputDir = GetSavePath();
            if (string.IsNullOrEmpty(outputDir))
            {
                Console.WriteLine("Operation cancelled.");
                return;
            }

            // Process each project
            Console.WriteLine($"\nProcessing {projects.Count} projects from the input file...");

            foreach (var project in projects)
            {
                try
                {
                    Console.WriteLine($"\nProcessing Project: {project.ProjectName}");
                    Console.WriteLine($"  Future Capex: ${project.FutureCapex:N2}");
                    Console.WriteLine($"  LOM Ounces: {project.LomOunces:N2}");
                    Console.WriteLine($"  Ounces Mined: {project.OuncesMined:N2}");

                    // Create project-specific output directory
                    string projectDir = Path.Combine(outputDir, project.ProjectName);
                    Directory.CreateDirectory(projectDir);

                    // Perform sensitivity analysis
                    SensitivityResults results = RunSensitivityAnalysis(
                        project.FutureCapex, project.LomOunces, project.OuncesMined);

                    // Save results to Excel
                    string excelPath = Path.Combine(projectDir, $"{project.ProjectName}_sensitivity_analysis.xlsx");
                    SaveToExcel(results, project, excelPath);
                    Console.WriteLine($"  Excel results saved to: {excelPath}");

                    // Create and save amortization rate heatmap
                    string amortHeatmapPath = Path.Combine(projectDir, $"{project.ProjectName}_amortization_sensitivity.png");
                    CreateHeatmap(
                        results.AmortizationRates,
                        results.PercentageLabels,
                        $"Amortization Rate Sensitivity Analysis: {project.ProjectName} ($/ounce)",
                        amortHeatmapPath);
                    Console.WriteLine($"  Amortization heatmap saved to: {amortHeatmapPath}");

                    // Create and save expected expense heatmap
                    string expenseHeatmapPath = Path.Combine(projectDir, $"{project.ProjectName}_expense_sensitivity.png");
                    CreateHeatmap(
                        results.ExpectedExpenses,
                        results.PercentageLabels,
                        $"Expected Expense Sensitivity: {project.ProjectName} ({project.OuncesMined:N0} Ounces Mined)",
                        expenseHeatmapPath);
                    Console.WriteLine($"  Expense heatmap saved to: {expenseHeatmapPath}");

                    Console.WriteLine($"  Sensitivity analysis completed for {project.ProjectName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing project {project.ProjectName}: {ex.Message}");
                }
            }

            Console.WriteLine("\nBatch processing completed!");
        }

        static string GetFilePath(string prompt)
        {
            while (true)
            {
                Console.Write(prompt);
                string filePath = Console.ReadLine().Trim();

                if (string.IsNullOrEmpty(filePath))
                {
                    return null;
                }

                if (File.Exists(filePath))
                {
                    return filePath;
                }
                else
                {
                    Console.WriteLine($"Error: File '{filePath}' does not exist.");
                }
            }
        }

        static string GetSavePath()
        {
            while (true)
            {
                Console.Write("\nEnter the folder path to save results (or press Enter to use current directory): ");
                string saveDir = Console.ReadLine().Trim();

                // Use current directory if input is empty
                if (string.IsNullOrEmpty(saveDir))
                {
                    return Directory.GetCurrentDirectory();
                }

                // Check if the directory exists
                if (Directory.Exists(saveDir))
                {
                    return saveDir;
                }
                else
                {
                    Console.Write($"Directory '{saveDir}' doesn't exist. Create it? (y/n): ");
                    string createDir = Console.ReadLine().ToLower();
                    if (createDir == "y")
                    {
                        try
                        {
                            Directory.CreateDirectory(saveDir);
                            Console.WriteLine($"Created directory: {saveDir}");
                            return saveDir;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error creating directory: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Please enter a valid directory path.");
                    }
                }
            }
        }

        static List<ProjectData> ReadInputExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // First worksheet
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Validate headers
                    string[] requiredColumns = { "Project", "Future_Capex", "LOM_Ounces", "Ounces_Mined" };
                    Dictionary<string, int> columnIndices = new Dictionary<string, int>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        string headerValue = worksheet.Cells[1, col].Text;
                        if (requiredColumns.Contains(headerValue))
                        {
                            columnIndices[headerValue] = col;
                        }
                    }

                    // Check if all required columns exist
                    var missingColumns = requiredColumns.Where(col => !columnIndices.ContainsKey(col)).ToList();
                    if (missingColumns.Any())
                    {
                        Console.WriteLine($"Error: Missing required columns in the Excel file: {string.Join(", ", missingColumns)}");
                        Console.WriteLine("Please ensure your Excel file has columns named: Project, Future_Capex, LOM_Ounces, Ounces_Mined");
                        return null;
                    }

                    // Read data rows
                    var projects = new List<ProjectData>();
                    for (int row = 2; row <= rowCount; row++) // Start from row 2 (skip header)
                    {
                        string projectName = worksheet.Cells[row, columnIndices["Project"]].Text;
                        
                        // Try to parse numeric values
                        double futureCapex, lomOunces, ouncesMined;
                        
                        if (!double.TryParse(worksheet.Cells[row, columnIndices["Future_Capex"]].Text, out futureCapex) ||
                            !double.TryParse(worksheet.Cells[row, columnIndices["LOM_Ounces"]].Text, out lomOunces) ||
                            !double.TryParse(worksheet.Cells[row, columnIndices["Ounces_Mined"]].Text, out ouncesMined))
                        {
                            Console.WriteLine($"Warning: Row {row} contains invalid numeric data. Skipping this project.");
                            continue;
                        }

                        if (string.IsNullOrEmpty(projectName))
                        {
                            Console.WriteLine($"Warning: Row {row} is missing a project name. Skipping this project.");
                            continue;
                        }

                        projects.Add(new ProjectData
                        {
                            ProjectName = projectName,
                            FutureCapex = futureCapex,
                            LomOunces = lomOunces,
                            OuncesMined = ouncesMined
                        });
                    }

                    if (projects.Count == 0)
                    {
                        Console.WriteLine("Error: No valid data rows found after cleaning.");
                        return null;
                    }

                    return projects;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
                return null;
            }
        }

        static SensitivityResults RunSensitivityAnalysis(double baseFutureCapex, double baseLomOunces, 
                                                        double ouncesMined, double variation = 0.20, int steps = 5)
        {
            // Calculate variation percentages
            int numSteps = (int)(2 * variation / (variation / steps)) + 1;
            double[] percentages = new double[numSteps];
            for (int i = 0; i < numSteps; i++)
            {
                percentages[i] = -variation + i * (variation / steps * 2);
            }

            // Create arrays for future capex and LOM ounces variations
            double[] futureCapexVariations = new double[percentages.Length];
            double[] lomOuncesVariations = new double[percentages.Length];
            for (int i = 0; i < percentages.Length; i++)
            {
                futureCapexVariations[i] = baseFutureCapex * (1 + percentages[i]);
                lomOuncesVariations[i] = baseLomOunces * (1 + percentages[i]);
            }

            // Create percentage labels for the table
            string[] percentageLabels = new string[percentages.Length];
            for (int i = 0; i < percentages.Length; i++)
            {
                percentageLabels[i] = $"{(int)(percentages[i] * 100)}%";
            }

            // Initialize matrices for the results
            double[,] amortMatrix = new double[percentages.Length, percentages.Length];
            double[,] expenseMatrix = new double[percentages.Length, percentages.Length];

            // Fill the matrices with amortization rates and expected expenses
            for (int i = 0; i < percentages.Length; i++)
            {
                for (int j = 0; j < percentages.Length; j++)
                {
                    // Calculate amortization rate
                    if (lomOuncesVariations[j] == 0)
                    {
                        amortMatrix[i, j] = double.NaN; // Avoid division by zero
                    }
                    else
                    {
                        amortMatrix[i, j] = futureCapexVariations[i] / lomOuncesVariations[j];
                    }

                    // Calculate expected expenses
                    expenseMatrix[i, j] = amortMatrix[i, j] * ouncesMined;
                }
            }

            return new SensitivityResults
            {
                AmortizationRates = amortMatrix,
                ExpectedExpenses = expenseMatrix,
                PercentageLabels = percentageLabels
            };
        }

        static void SaveToExcel(SensitivityResults results, ProjectData project, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                // Add Amortization Rates worksheet
                var amortSheet = package.Workbook.Worksheets.Add("Amortization Rates");
                
                // Add row and column headers
                for (int i = 0; i < results.PercentageLabels.Length; i++)
                {
                    amortSheet.Cells[1, i + 2].Value = results.PercentageLabels[i];
                    amortSheet.Cells[i + 2, 1].Value = results.PercentageLabels[i];
                }
                
                // Add header titles
                amortSheet.Cells[1, 1].Value = "Future Capex / LOM Ounces";
                
                // Add data
                for (int i = 0; i < results.PercentageLabels.Length; i++)
                {
                    for (int j = 0; j < results.PercentageLabels.Length; j++)
                    {
                        amortSheet.Cells[i + 2, j + 2].Value = results.AmortizationRates[i, j];
                        amortSheet.Cells[i + 2, j + 2].Style.Numberformat.Format = "0.00";
                    }
                }
                
                // Format headers
                using (var range = amortSheet.Cells[1, 1, 1, results.PercentageLabels.Length + 1])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                
                using (var range = amortSheet.Cells[1, 1, results.PercentageLabels.Length + 1, 1])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                
                amortSheet.Cells.AutoFitColumns();

                // Add Expected Expenses worksheet
                var expenseSheet = package.Workbook.Worksheets.Add("Expected Expenses");
                
                // Add row and column headers
                for (int i = 0; i < results.PercentageLabels.Length; i++)
                {
                    expenseSheet.Cells[1, i + 2].Value = results.PercentageLabels[i];
                    expenseSheet.Cells[i + 2, 1].Value = results.PercentageLabels[i];
                }
                
                // Add header titles
                expenseSheet.Cells[1, 1].Value = "Future Capex / LOM Ounces";
                
                // Add data
                for (int i = 0; i < results.PercentageLabels.Length; i++)
                {
                    for (int j = 0; j < results.PercentageLabels.Length; j++)
                    {
                        expenseSheet.Cells[i + 2, j + 2].Value = results.ExpectedExpenses[i, j];
                        expenseSheet.Cells[i + 2, j + 2].Style.Numberformat.Format = "0.00";
                    }
                }
                
                // Format headers
                using (var range = expenseSheet.Cells[1, 1, 1, results.PercentageLabels.Length + 1])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                
                using (var range = expenseSheet.Cells[1, 1, results.PercentageLabels.Length + 1, 1])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                
                expenseSheet.Cells.AutoFitColumns();

                // Add Input Summary worksheet
                var summarySheet = package.Workbook.Worksheets.Add("Input Summary");
                
                summarySheet.Cells[1, 1].Value = "Parameter";
                summarySheet.Cells[1, 2].Value = "Value";
                
                summarySheet.Cells[2, 1].Value = "Project";
                summarySheet.Cells[2, 2].Value = project.ProjectName;
                
                summarySheet.Cells[3, 1].Value = "Future Capex";
                summarySheet.Cells[3, 2].Value = project.FutureCapex;
                summarySheet.Cells[3, 2].Style.Numberformat.Format = "#,##0.00";
                
                summarySheet.Cells[4, 1].Value = "LOM Ounces";
                summarySheet.Cells[4, 2].Value = project.LomOunces;
                summarySheet.Cells[4, 2].Style.Numberformat.Format = "#,##0.00";
                
                summarySheet.Cells[5, 1].Value = "Ounces Mined";
                summarySheet.Cells[5, 2].Value = project.OuncesMined;
                summarySheet.Cells[5, 2].Style.Numberformat.Format = "#,##0.00";
                
                // Format headers
                using (var range = summarySheet.Cells[1, 1, 1, 2])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                
                summarySheet.Cells.AutoFitColumns();

                // Save the Excel file
                package.SaveAs(new FileInfo(filePath));
            }
        }

        static void CreateHeatmap(double[,] data, string[] labels, string title, string filePath)
        {
            var plt = new ScottPlot.Plot(800, 600);

            // Create heatmap
            var heatmap = plt.AddHeatmap(data, lockScales: false);
            heatmap.ShowAxisLabels = true;
            heatmap.XTickLabels = labels;
            heatmap.YTickLabels = labels;
            
            // Add a colorbar
            plt.AddColorbar(heatmap);

            // Set plot properties
            plt.Title(title);
            plt.XLabel("LOM Ounces Variation");
            plt.YLabel("Future Capex Variation");
            
            // Customize axis ticks
            plt.XAxis.TickLabelStyle(rotation: 45);
            
            // Display values in cells
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    // Invert the i-coordinate for plotting due to how heatmaps are displayed
                    string text = $"{data[i, j]:F2}";
                    plt.AddText(text, j, i, size: 12, 
                                color: data[i, j] > (data.Cast<double>().Average()) ? Color.White : Color.Black);
                }
            }

            // Save to file
            plt.SaveFig(filePath);
        }
    }

    class ProjectData
    {
        public string ProjectName { get; set; }
        public double FutureCapex { get; set; }
        public double LomOunces { get; set; }
        public double OuncesMined { get; set; }
    }

    class SensitivityResults
    {
        public double[,] AmortizationRates { get; set; }
        public double[,] ExpectedExpenses { get; set; }
        public string[] PercentageLabels { get; set; }
    }
}
