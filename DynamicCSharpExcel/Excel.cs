using System;

namespace DynamicCsharp
{
    /// <summary>
    /// https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
    /// 
    /// Additional enhancements are possible when you call a COM type that does not require a primary interop assembly (PIA) at run time. Removing the dependency on PIAs results in version independence and easier deployment. For more information about the advantages of programming without PIAs, see Walkthrough: Embedding Types from Managed Assemblies.
    ///
    ///In addition, programming is easier because the types that are required and returned by COM methods can be represented by using the type dynamic instead of Object.Variables that have type dynamic are not evaluated until run time, which eliminates the need for explicit casting.For more information, see Using Type dynamic.
    ///
    ///In C# 4, embedding type information instead of using PIAs is default behavior. Because of that default, several of the previous examples are simplified because explicit casting is not required.
    /// </summary>


    public class Excel
    {
        private readonly ExcelOptions _excelOptions;

        private dynamic _excelObj;
        private dynamic _workbook;

        public dynamic DefaultWorksheet;

        /// <summary>
        /// Version information, use for logging purposes in case of errors
        /// </summary>
        public string Version => $"{_excelObj.Value} Version {_excelObj.Version} Build {_excelObj.Build}";

        public Excel(ExcelOptions excelOptions)
        {
            _excelOptions = excelOptions;
            CreateExcelObj();
            CreateWorksheet(_excelOptions.WorksheetName);
        }

        private void CreateExcelObj()
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application", true);
            _excelObj = Activator.CreateInstance(excelType);
            _excelObj.Visible = _excelOptions.Visible;

            // Setting to false so that number values in cell can be text and warning doesn't show
            _excelObj.ErrorCheckingOptions.NumberAsText = false;
        }

        private void CreateWorksheet(string worksheetName)
        {
            _workbook = _excelObj.Workbooks.Add();
            DefaultWorksheet = _excelObj.ActiveSheet;

            if (!string.IsNullOrEmpty(worksheetName))
            {
                DefaultWorksheet.Name = worksheetName;
            }
        }

        public void AutofitColumn(int col)
        {
            DefaultWorksheet.Columns[col].Autofit();
        }

        public void WriteCellValue(int row, int col, object value)
        {
            DefaultWorksheet.Cells[row, col] = value;
        }

        public void FormatCell(int row, int col, string format, bool isBold)
        {
            DefaultWorksheet.Cells[row, col].Font.Bold = isBold;
            if (!string.IsNullOrEmpty(format))
            {
                DefaultWorksheet.UsedRange.Cells[row, col].NumberFormat = format;
            }
        }

        public void SetEntireColumnFormat(int col, string format)
        {
            DefaultWorksheet.UsedRange.Cells[1, col].EntireColumn.NumberFormat = format;
        }

        public void SetEntireColumnStyle(int col, string style)
        {
            DefaultWorksheet.UsedRange.Cells[1, col].EntireColumn.Style = style;
        }

        public void SetEntireRowStyle(int row, string style)
        {
            DefaultWorksheet.UsedRange.Cells[1, row].EntireRow.Style = style;
        }

        private void ShowAvailableStyles()
        {

            // Where all available styles in a workbook is stored.
            dynamic styles = _workbook.Styles;

            foreach (var style in styles)
            {
                Console.WriteLine($"{style.Name} {style.BuiltIn}");
            }
        }

        public void SaveAndQuit()
        {
            Save();
            Quit();
        }

        public void Save()
        {
            _workbook.SaveAs(_excelOptions.Filename);
        }

        public void Quit()
        {
            _excelObj.Workbooks.Close();
            _excelObj.Quit();
        }
    }
}