using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using applications = System.Windows.Forms.Application;
using Microsoft.CSharp.RuntimeBinder;
/// <summary>
/// Excelling is suppose to deal with all of the information that leads into Microsoft Excel relating to the external database that the addon application uses
/// Author: Millar Sue msue619
/// 2018 p4p #122: Risk information management in Building Information Modelling (BIM)
/// </summary>
namespace TrialsWithForm
{

    public static class Excels
    {
        public static Excel.Application excel = null;
        public static Excel.Workbook Database = null;
        //WriteExcel allows the new excel that is made by the addon application to be formatted in the way that works with the add on application
        public static Excel.Application WriteExcel()
        {
            if (excel == null)
            {
                excel = new Excel.Application();
            }
            excel.Visible = true;
            Database = excel.Workbooks.Add();
            var actsheet = (Excel.Worksheet)Database.ActiveSheet;
            actsheet.Columns["A:G"].ColumnWidth = 17;
            List<string> headings = new List<string>(new string[] { "EIDs", "Risk Type", "Name", "Discipline", "Description", "Likelihood", "Severity" });
            int rows = 1, cols = 1;

            for (int i = 0; i < headings.Count(); i++)
            {
                actsheet.Cells[rows, cols] = headings[i];
                cols++;
            }
            Excel.Range formatRange;
            formatRange = actsheet.get_Range("A:A");
            formatRange.NumberFormat = "@";
            return excel;
        }
        //OpenDatabase opens the excel database that is mentioned in the string fileName.
        //fileName represents the file path directory
        public static void OpenDatabase(string fileName)
        {
            excel = new Excel.Application();
            Database = excel.Workbooks.Open(fileName);
            var actsheet = (Excel.Worksheet)Database.ActiveSheet;
            excel.Visible = true;

        }
        //SaveDatabase saves the active database that is within the addon application memory
        public static void SaveDatabase()
        {
            Database = excel.ActiveWorkbook;
            Database.Save();
        }
        //SaveDatabase closes the active database that is within the addon application memory
        public static void CloseDatabase(Excel.Application Cexcel)
        {
            try
            {
                Database = Cexcel.ActiveWorkbook;
                Database.Close();
                Cexcel.Quit();
            }
            catch { }
        }

        //DataIntoExcel places the relevant information in the addon "InputInformation" tab into the excel database
        public static void DataIntoExcel(string eids, string risktype, string title, string discipline, string description, string severity, string likelihood)
        {
            excel = WindowsFormsApp3.TabulatedForms.GetExcel();
            Database = excel.ActiveWorkbook;
            Excel.Worksheet actsheet = Database.ActiveSheet;
            List<string> headings = new List<string>(new string[] { eids, risktype, title, discipline, description,severity, likelihood });


            var usedRows = excel.WorksheetFunction.CountA(actsheet.Columns[3]);
            int cols = 1;

            for (int i = 0; i < headings.Count(); i++)
            {
                actsheet.Cells[usedRows + 1, cols] = headings[i];
                cols++;
            }
        }
        //CountA works similar to the "CountA" command in Microsoft Excel Application which counts all active cells that has information
        // in the Severity column
        public static int countA(Excel.Worksheet xlSheet)
        {
            int flag = 0;
            List<int> trials = new List<int>();
            int i = 2;
            while (flag != 1)
            {
                var num1 = xlSheet.Cells[i, 6].Value;

                try
                {
                    int num2 = (int)num1;
                    trials.Add(num2);
                    i++;
                }
                catch
                {
                    flag = 1;
                }
            }
            return (i - 1);
        }

    }
}
