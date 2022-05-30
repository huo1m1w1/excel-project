using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                // change to your own directory
                workbook = app.Workbooks.Open("C:\\Users\\huomi\\OneDrive\\My Documents\\MScFE 670\\GWP 1\\WorldQuant_Module3_CSA_SkeletonCode\\WorldQuant_Module3_CSA_SkeletonCode\\property_pricing.xlsx", ReadOnly: false);
            }
            catch
            {
                SetUp();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                } catch { }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            // TODO: Implement this method
            app.Workbooks.Add();
            workbook = app.ActiveWorkbook;
            workbook.Worksheets.Add();

            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            currentSheet.Name = "property_pricing";

            currentSheet.Cells[1, "A"] = "Size (in square feet)";
            currentSheet.Cells[1, "B"] = "Suburb";
            currentSheet.Cells[1, "C"] = "City";
            currentSheet.Cells[1, "D"] = "Market Value";

            // change it to your own directory, and there is a file in the folder already, which can be deleted to reenter a new file
            workbook.SaveAs("C:\\Users\\huomi\\OneDrive\\My Documents\\MScFE 670\\GWP 1\\WorldQuant_Module3_CSA_SkeletonCode\\WorldQuant_Module3_CSA_SkeletonCode\\property_pricing.xlsx");


        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            // TODO: Implement this method
            int row = 2;
            Excel.Worksheet propertyprice = workbook.Worksheets[1];

            while (true)
            {
                if (propertyprice.Cells[row, "A"].Value ==null)
                {
                    propertyprice.Cells[row, "A"].Value = size;
                    propertyprice.Cells[row, "B"].Value = suburb;                    
                    propertyprice.Cells[row, "C"].Value = city;
                    propertyprice.Cells[row, "D"].Value = value;
                    break;

                }

                row++;
            }
            
            
        }

        static float CalculateMean()
        {
            // TODO: Implement this method            

            int row = 2;
            Excel.Worksheet currentsheet = workbook.Worksheets[1];

            while (true)
            {
                if (currentsheet.Cells[row, "A"].Value == null)
                {
                    string end = "D" + row;
                    int a = row + 2;
                    string loc = "D" + a;
                    currentsheet.Range[loc].Formula = "=AVERAGE(D2:" + end + ")";                    
                    
                    currentsheet.Cells[row + 2, 3] = "Mean";
                    // change to you own directory
                    workbook.SaveAs("C:\\Users\\huomi\\OneDrive\\My Documents\\MScFE 670\\GWP 1\\WorldQuant_Module3_CSA_SkeletonCode\\WorldQuant_Module3_CSA_SkeletonCode\\property_pricing.xlsx");

                   

                    return currentsheet.Cells[row + 2, 3];
                    break;
                }
                row++;                
            }
            
        }

        static float CalculateVariance()
        {
            // TODO: Implement this method
            int row = 2;
            Excel.Worksheet currentsheet = workbook.Worksheets[1];

            while (true)
            {
                if (currentsheet.Cells[row, "A"].Value == null)
                {
                    string end = "D" + row;
                    int a = row + 3;
                    string loc = "D" + a;
                    currentsheet.Range[loc].Formula = "=VAR.P(D2:" + end + ")";

                    currentsheet.Cells[a, 3] = "VANRIANCE";
                    return currentsheet.Cells[row + 3, 3];

                    break;
                }
                row++;
            }
            workbook.SaveAs("C:\\Users\\huomi\\WorldQuant_Module3_CSA_SkeletonCode\\WorldQuant_Module3_CSA_SkeletonCode\\property_pricing.xlsx");

            
        }

        static float CalculateMinimum()
        {
            // TODO: Implement this method
            int row = 2;
            Excel.Worksheet currentsheet = workbook.Worksheets[1];

            while (true)
            {
                
                if (currentsheet.Cells[row, "A"].Value == null)
                {
                    string end = "D" + row;
                    int a = row + 4;
                    string loc = "D" + a;
                    currentsheet.Range[loc].Formula = "=MIN(D2:" + end + ")";

                    currentsheet.Cells[a, 3] = "Min";
                    workbook.SaveAs("C:\\Users\\huomi\\WorldQuant_Module3_CSA_SkeletonCode\\WorldQuant_Module3_CSA_SkeletonCode\\property_pricing.xlsx");
                    return currentsheet.Cells[row + 4, 3];
                    
                }
                row++;
            }
            
            
        }

        static float CalculateMaximum()
        {
            // TODO: Implement this method
            int row = 2;
            Excel.Worksheet currentsheet = workbook.Worksheets[1];

            while (true)
            {
                if (currentsheet.Cells[row, "A"].Value == null)
                {
                    string end = "D" + row;
                    int a = row + 5;
                    string loc = "D" + a;
                    currentsheet.Range[loc].Formula = "=MAX(D2:" + end + ")";

                    currentsheet.Cells[a, 3] = "Max";
                    workbook.SaveAs("C:\\Users\\huomi\\WorldQuant_Module3_CSA_SkeletonCode\\WorldQuant_Module3_CSA_SkeletonCode\\property_pricing.xlsx");
                    return currentsheet.Cells[row + 5, 3];
                    break;
                }
                row++;
            }

        
            
        }
    }
}
