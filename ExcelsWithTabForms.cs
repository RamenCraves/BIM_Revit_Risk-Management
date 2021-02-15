using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using Excel = Microsoft.Office.Interop.Excel;
using cl = Microsoft.Office.Interop.Excel.Range;

/// <summary>
/// TrialsWithTabForms deals with majority of the commands that relate with the RIM operability
/// Author: Millar Sue msue619
/// 2018 p4p #122: Risk information management in Building Information Modelling (BIM)
/// </summary>
/// 
namespace TrialsWithTabForms
{

    public partial class ExcelsWithTabsForms
    {
        public static List<int> el;
        public static GroupBox[][] rim;
        public static int linkIndex;
        // ReadAndWriteRM reads the excel database and inputs the necessary information into the RIM risk matrix.
        public static void ReadAndWriteRM(Excel.Workbook Database, GroupBox[][] textBoxes)
        {
            rim = textBoxes;
            Excel.Worksheet xlSheet = Database.ActiveSheet;
            int activeRows = TrialsWithForm.Excels.countA(xlSheet); int y; int j; int k; int x = 10; int counts;
            for (int iem = 1; iem < activeRows; iem++)
            {

                j = (int)xlSheet.Cells[iem + 1, 6].Value - 1;
                k = (int)xlSheet.Cells[iem + 1, 7].Value - 1;
                string title = xlSheet.Cells[iem + 1, 3].Value;
                counts = textBoxes[j][k].Controls.Count;
                y = 10 + counts * 25;
                CreateLink(title, string.Concat(" ", iem), x, y, j, k, textBoxes);

            }
        
        }
        //CreateLink creates the link that is placed in the Risk matrix as depicted by the first column of the excel database.
        public static void CreateLink(string text, string name, int x, int y, int j, int k, GroupBox[][] textBoxes)
        {
            LinkLabel link = new LinkLabel();
            link.Text = text;
            link.Name = name;
            link.Font = new Font("Times New Roman", 9);
            link.AutoSize = true;

            link.Location = new System.Drawing.Point(x, y);
            link.ForeColor = System.Drawing.Color.Blue;
            link.BorderStyle = BorderStyle.None;
            textBoxes[j][k].Controls.Add(link);

            link.LinkClicked += new LinkLabelLinkClickedEventHandler(linkClick);
        }
        //linkClick iworks in conjunction with the CreateLink command and is used to create the necessary commands that are used in th  event th  link is clicked
        public static void linkClick(object sender, LinkLabelLinkClickedEventArgs e)

        {
            string linkstxt;
            linkstxt = ((LinkLabel)sender).Text;
            Excel.Application excel = WindowsFormsApp3.TabulatedForms.GetExcel();
            Excel.Workbook Database = excel.ActiveWorkbook;
            Excel.Worksheet xlSheet = Database.ActiveSheet;
            int activeRows = TrialsWithForm.Excels.countA(xlSheet);
            var IDs = GetIDs(linkstxt, activeRows, xlSheet);
            var curr_color = GetColor(linkstxt, activeRows, xlSheet);
            el = getBetweenints(IDs, " "," ");
            Highlights(el);
            

        }
        //Highlights highlights the elements that are needed to be highlighted as a result of the event of a link label being pressed
        public static void Highlights(List<int> eltrials)
        {

            Execute(WindowsFormsApp3.TabulatedForms.commandData,eltrials);
        }

        public static string GetIDs(string title, int activeRows, Excel.Worksheet xlSheet)
        {
            string returnID = "Cannot find Title"; string Focus;
            for (int i = 2; i <= activeRows; i++)
            {
                Focus = xlSheet.Cells[i, 3].Value;
                if (Focus == title)
                {
                    returnID = xlSheet.Cells[i, 1].Value;
                    WindowsFormsApp3.TabulatedForms.ChangeReRiskDes(i);
                }
            }
            return returnID;
        }

        // Result Execute highlights the elements whose element id is within the eletrials list of ints.
        public static Result Execute(ExternalCommandData eData, List<int> eletrials)
        {
            UIApplication uiapp = eData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();

            View3D view3d = doc.ActiveView as View3D;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Transaction Name");
                ICollection<ElementId> ids = new List<ElementId>();
                foreach (int element in eletrials)
                {
                    ElementId id = new ElementId(element);
                    Element eleFromId = doc.GetElement(id);
                    ids.Add(eleFromId.Id);

                }
                uidoc.Selection.SetElementIds(ids);

                uidoc.ShowElements(ids);



                tx.Commit();
            }
            return Result.Succeeded;
        }
        //A string with the relevant information that is to be described as int is passed into getBetweenints, whereby the characters for the start and end is specified. The index of each of the start and the end 
        // constraints are taken and the string between these two indexes is taken, converted into an integer and the string that ws passed is chopped so that this process could be repeated til the end of the string.
        // The integers that are in between the two constraints are placed in a list of ints and after the process is finished is pased back out of the command
        public static List<int> getBetweenints(string pass, string startConstraint, string endConstraint)
        {
            List<int> ids = new List<int>();
            int count = 0; int length = pass.Length; string num; int numInt = 0;int flag; 
            int Start, End;

            while (count + 1 != pass.Length)
            {
                Start = pass.IndexOf(startConstraint, count) + 1;
                End = pass.IndexOf(endConstraint, Start);
                num = pass.Substring(Start, End - Start); 
                count = End;
                numInt = Convert.ToInt32(num);
                flag = 0;
                foreach (int i in ids)
                {
                    if (i == numInt)
                    { flag = 1; }
                }
                if ( flag ==0)
                {
                    ids.Add(numInt);
                }


            }
            return ids;
        }
        //A string with the relevant information that is to be described as string is passed into getBetweendtring, whereby the characters for the start and end is specified. The index of each of the start and the end 
        // constraints are taken and the string between these two indexes is taken and the string that ws passed is chopped so that this process could be repeated til the end of the string.
        // The strings that are in between the two constraints are placed in a list of string and after the process is finished is pased back out of the command
        public static List<string> getBetweenstrings(string pass, string startConstraint, string endConstraint)
        {
            List<string> ids = new List<string>();
            int count = 0; int length = pass.Length; string num; int flag;
            int Start, End;

            while (count + 1 != pass.Length)
            {
                Start = pass.IndexOf(startConstraint, count) + 1;
                End = pass.IndexOf(endConstraint, Start);
                num = pass.Substring(Start, End - Start);
                count = End;
                flag = 0;
                foreach (string i in ids)
                {
                    if (i == num)
                    { flag = 1; }
                }
                if (flag == 0)
                {
                    ids.Add(num);
                }


            }
            return ids;
        }
        // Gets the back colour of the Risk matrix coresponding to the risk severity and the risk discipline.
        public static List<int> GetColor(string title, int activeRows, Excel.Worksheet xlSheet)
        {
            string Focus; System.Drawing.Color true_color = System.Drawing.Color.Blue;
            List<int> colour = new List<int>();
            for (int i = 2; i <= activeRows; i++)
            {
                Focus = xlSheet.Cells[i, 3].Value;
                if (Focus == title)
                {
                    var severity = (int)xlSheet.Cells[i, 6].Value;
                    var likelihood = (int)xlSheet.Cells[i, 7].Value;
                    int blues = rim[severity-1][likelihood-1].BackColor.B;
                    int greens = rim[severity-1][likelihood-1].BackColor.G;
                    int reds = rim[severity-1][likelihood-1].BackColor.R;
                    colour.Add(reds); colour.Add(greens); colour.Add(blues);
                }
            }
            return colour;
        }
    }
}

