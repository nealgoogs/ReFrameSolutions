using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string filePath = "F:\\ReFrameSolutions\\James Smith Transactions.xlsx";  // Update this with your actual file path

        // Initialize tracking variables
        double totalMoneyInBank = 0;
        double totalDonations = 0;
        double totalExpenses = 0;
        double totalDebt = 0;

        // Read the Excel file
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Assuming first row has headers
            {
                string transactionType = worksheet.Cells[row, 1].Value.ToString().Trim();
                double transactionAmount = Convert.ToDouble(worksheet.Cells[row, 2].Value);

                switch (transactionType)
                {
                    case "Montary Donation":
                        totalMoneyInBank += transactionAmount;
                        totalDonations += transactionAmount;
                        break;

                    case "Non-monetary Donation":
                        totalDonations += transactionAmount;
                        break;

                    case "Monetary Expense":
                        totalMoneyInBank -= transactionAmount;
                        totalExpenses += transactionAmount;
                        break;

                    case "Non-monetary Expense":
                        totalExpenses += transactionAmount;
                        break;

                    case "Loan":
                        totalMoneyInBank += transactionAmount;
                        totalDebt += transactionAmount;
                        break;

                    case "Loan Payment":
                        totalMoneyInBank -= transactionAmount;
                        totalDebt -= transactionAmount;
                        break;

                    default:
                        Console.WriteLine($"Unknown transaction type: {transactionType}");
                        break;
                }
            }
        }

        Console.WriteLine("===============================================");
        Console.WriteLine("             TRANSACTION SUMMARY               ");
        Console.WriteLine("===============================================");

        // Table-like formatting
        Console.WriteLine("{0,-30} {1,15:C2}", "Total Money in the Bank:", totalMoneyInBank);
        Console.WriteLine("{0,-30} {1,15:C2}", "Total Donations:", totalDonations);
        Console.WriteLine("{0,-30} {1,15:C2}", "Total Expenses:", totalExpenses);
        Console.WriteLine("{0,-30} {1,15:C2}", "Remaining Debt:", totalDebt);

        Console.WriteLine("===============================================");
        Console.WriteLine("    All calculations completed successfully!"   );
        Console.WriteLine("===============================================");
    }
}
