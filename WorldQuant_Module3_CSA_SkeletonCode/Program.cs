using System;
using Excel = Microsoft.Office.Interop.Excel;

/* WQU Group 1A */

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
                workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
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
            workbook = app.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            worksheet.Name = "Property";

            worksheet.Cells[1, 1] = "Size";
            worksheet.Cells[1, 2] = "Suburb";
            worksheet.Cells[1, 3] = "City";
            worksheet.Cells[1, 4] = "Market value";
            worksheet.Cells[1, 5] = 0;

            workbook.SaveAs("property_pricing.xlsx");
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            int row = (int)worksheet.Cells[1, 5].Value;
            row++;

            worksheet.Cells[row + 1, 1] = size;
            worksheet.Cells[row + 1, 2] = suburb;
            worksheet.Cells[row + 1, 3] = city;
            worksheet.Cells[row + 1, 4] = value;

            worksheet.Cells[1, 5] = row;
        }

        static float CalculateMean()
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            int count = (int)worksheet.Cells[1, 5].Value;

            float sum = 0.0f;
            for (int i = 1; i <= count; i++)
            {
                sum += worksheet.Cells[i + 1, 4].Value;
            }

            return sum / count;
        }

        static float CalculateVariance()
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            int count = (int)worksheet.Cells[1, 5].Value;
            float mean = CalculateMean();

            float variance = 0.0f;
            for (int i = 1; i <= count; i++)
            {
                float value = (float)worksheet.Cells[i + 1, 4].Value;
                variance += (float)Math.Pow((value - mean), 2);
            }

            return variance / count;
        }

        static float CalculateMinimum()
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            int count = (int)worksheet.Cells[1, 5].Value;

            float min = (float)worksheet.Cells[2, 4].Value;
            for (int i = 2; i <= count; i++)
            {
                float value = (float)worksheet.Cells[i + 1, 4].Value;
                if (value < min)
                {
                    min = value;
                }
            }

            return min;
        }

        static float CalculateMaximum()
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            int count = (int)worksheet.Cells[1, 5].Value;

            float max = (float)worksheet.Cells[2, 4].Value;
            for (int i = 2; i <= count; i++)
            {
                float value = (float)worksheet.Cells[i + 1, 4].Value;
                if (value > max)
                {
                    max = value;
                }
            }

            return max;
        }
    }
}
