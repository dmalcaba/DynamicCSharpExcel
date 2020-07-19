using DynamicCsharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using DynamicCSharpOutlook;
using DynamicCSharpWord;

namespace DynamicCSharpConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //Run();
            //SendOutlookEmail();
            //CreateWord();
            new ReflectionsReview().Start();

            Console.WriteLine("\n\nPress enter to continue...");
            Console.ReadLine();
        }

        private static void FirstTest()
        {
            var myExcel = new Excel(new ExcelOptions {Filename = "TestExcel", Visible = true, WorksheetName = "My Work"});

            // For logging purpose
            Console.WriteLine($"Excel build: {myExcel.Version}");

            myExcel.WriteCellValue(1, 1, "This is the name column");
            myExcel.WriteCellValue(1, 2, "This should go in B1; quick brown fox");
            myExcel.WriteCellValue(3, 2, "should be in b3");
            myExcel.WriteCellValue(3, 3, "good gracious great balls of fire!!!");
            // Number
            myExcel.WriteCellValue(3, 4, 12345.7890);
            myExcel.FormatCell(3, 4, "$ #,##0.00", false);
            // Date
            myExcel.WriteCellValue(3, 5, DateTime.Today);
            myExcel.FormatCell(3, 5, "dd-MMM-yyyy", false);
            myExcel.FormatCell(3, 6, string.Empty, true);


            //custom format
            var format = "[Red][<0]$ -#,##0.00; [Black][>0]$ #,##0.00";

            myExcel.WriteCellValue(4, 6, -345.346);
            myExcel.FormatCell(4, 6, format, false);

            myExcel.WriteCellValue(5, 6, 34545.346);
            myExcel.FormatCell(5, 6, format, false);

            myExcel.AutofitColumn(1);
            myExcel.AutofitColumn(2);
            myExcel.AutofitColumn(3);

            for (int i = 1; i < 21; i++)
            {
                myExcel.WriteCellValue(i, 7, new Random().Next());
            }

            myExcel.AutofitColumn(7);

            myExcel.SetEntireColumnFormat(7, "$ #,##0.00");
            myExcel.SetEntireColumnStyle(7, "Currency");
            
            myExcel.SaveAndQuit();
        }

        public static void Run()
        {
            var myExcel = new Excel(new ExcelOptions {Filename = "TestExcel", Visible = true, WorksheetName = "My Work"});

            Debug.WriteLine($"Version: {myExcel.Version}");

            var accountList = new List<AccountValues>
            {
                new AccountValues("11100010", "Petty Cash", "001", 10.80m),
                new AccountValues("11100026", "Bank - Cemetery Boards", "201", -27678.05m),
            };

            myExcel.WriteCellValue(1, 1, "Account Code");
            myExcel.WriteCellValue(1, 2, "Account Name");
            myExcel.WriteCellValue(1, 3, "Cost Center");
            myExcel.WriteCellValue(1, 4, "Amount");
            
            myExcel.SetEntireRowStyle(1, ExcelStyle.Heading);

            int row = 1;

            // Set format as Text  - so that Numbers will be stored as Text and values like 001 will not be stored as 1
            myExcel.SetEntireColumnFormat(1, "@");
            myExcel.SetEntireColumnFormat(3, "@");

            foreach (var accountValues in accountList)
            {
                row++;
                myExcel.WriteCellValue(row, 1, accountValues.AccountCode);
                myExcel.WriteCellValue(row, 2, accountValues.AccountName);
                myExcel.WriteCellValue(row, 3, accountValues.CostCenter);
                myExcel.WriteCellValue(row, 4, accountValues.Amount);
            }

            myExcel.SetEntireColumnStyle(4, ExcelStyle.Currency);
            myExcel.DefaultWorksheet.UsedRange.Columns.Autofit();
        }

        public static void CreateWord()
        {
            var wordObj = new Word();

            // For logging purpose
            Debug.WriteLine($"Information: {wordObj.Version}");

        }

        public static void SendOutlookEmail()
        {
            var options = new OutlookOptions
            {
                Recipients = new List<string> {"someone@github.com"},
                Subject = "Test email using C# code",
                Body = "<quote>This is a quote</quote>"
            };

            //options.Attachments.Add(@"C:\Users\Public\Documents\ViewResultExtension.txt");

            var outlook = new Outlook(options);

            // For logging purpose
            Debug.WriteLine($"Outlook build: {outlook.Version}");

            //outlook.SendMail();
        }
    }

}
