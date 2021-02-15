using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Excel = Microsoft.Office.Interop.Excel;
using applications = System.Windows.Forms.Application;
/// <summary>
/// TabulatedForms is the addon portal connecting and providing the "Active link" feature to the Parent application: Autodesk Revit and the External database: Microsoft Excel
/// Author: Millar Sue msue619
/// 2018 p4p #122: Risk information management in Building Information Modelling (BIM)
/// </summary>
namespace WindowsFormsApp3
{

    public partial class TabulatedForms : System.Windows.Forms.Form
    {
        public string files = "";
        public string member = " ";
        public XYZ xyz;
        public static string IDs;
        public static List<int> el;
        public static Excel.Application excel = null;
        public static Excel.Workbook Database = null;
        public static ExternalCommandData commandData;
        public System.Windows.Forms.TextBox riskPointID = new System.Windows.Forms.TextBox();
        public static System.Windows.Forms.TextBox ReRiskDes;
        public string fsFamilyName;
        public string fsName;
        public string severity; public string likelihood;
        public static string reriskmessage = "";

        //CreateTabWithForms just runs the application and if the application is already opened then it does nothing because of the try catch statement
        public static void CreateTabWithForms(ExternalCommandData commandData)
        {
            try
            {
                applications.EnableVisualStyles();
                applications.Run(new WindowsFormsApp3.TabulatedForms(commandData));
            }
            catch
            {
            }

        }
        //TabulatedForms runs the addon GUI application and sets up the visuals. 
        //An important thing to note here is that the commandData that is stored for the active application data is taken out of the application to become a global variable.
        public TabulatedForms(ExternalCommandData cData)
        {
            commandData = cData;
            InitializeComponent();
            tabControl.TabPages.Remove(pointControl);
            UIApplication app = commandData.Application;
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            ReRiskDes = ReRisk;
        }
        //OK_button_Click is activated when the OK button is clicked on the "InputInformation tab" which allows all the risk information to be taken from the GUI and placed in the database that is loaded.
        public void OK_button_Click(object sender, EventArgs e)
        {
            if (S1.Checked) { severity = S1.Text; }
            else if (S2.Checked) { severity = S2.Text; }
            else if (S3.Checked) { severity = S3.Text; }
            else if (S4.Checked) { severity = S4.Text; }
            else if (S5.Checked) { severity = S5.Text; }

            if (L1.Checked) { likelihood = L1.Text; }
            else if (L2.Checked) { likelihood = L2.Text; }
            else if (L3.Checked) { likelihood = L3.Text; }
            else if (L4.Checked) { likelihood = L4.Text; }
            else if (L5.Checked) { likelihood = L5.Text; }
            if (excel != null)
            { 
                if (severity == " ") { MessageBox.Show("Please fill in the risk severity level"); }
                if (likelihood == " ") { MessageBox.Show("Please fill in the risk likelihood level"); }

                TrialsWithForm.Excels.DataIntoExcel(elementBox.Text, typeBox.Text, RiskTitle.Text, RiskDisci.Text, RiskDes.Text, severity, likelihood);
                TrialsWithForm.Excels.SaveDatabase();
                GroupBox[][] textBoxes = getGroupBoxes();
                ClearContents();
                TrialsWithTabForms.ExcelsWithTabsForms.ReadAndWriteRM(Database, textBoxes);
            }
            else
            {
                MessageBox.Show("No Database in memory");
            }

        }
        //CloseDatabase closes the active database that is within the application's memory
        private void CloseDatabase_Click(object sender, EventArgs e)
        {
            Excel.Application excels = GetExcel();
            TrialsWithForm.Excels.CloseDatabase(excels);
            ClearContents();
        }
        //OpenDatabase opens a database to be loaded into the application's memory
        public void OpenDatabase_Click(object sender, EventArgs e)
        {
            string fileName = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = "C:\\Users\\mini_\\Desktop";
                openFileDialog1.Filter = "All Excel Files (*.xlsx*)|*.xls*";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    excel = new Excel.Application();
                    Database = excel.Workbooks.Open(fileName);
                    var actsheet = (Excel.Worksheet)Database.ActiveSheet;
                    excel.Visible = true;
                }
                files = fileName;
            }

            typeBox.ReadOnly = true;
            elementBox.ReadOnly = true;
            GroupBox[][] textBoxes = getGroupBoxes();
            TrialsWithTabForms.ExcelsWithTabsForms.ReadAndWriteRM(Database, textBoxes);

        }
        //NewData_Click allows the formation of a new database and to be formated in the correct way to not cause the addon application to be throw an exception in its runtime.
        private void NewData_Click(object sender, EventArgs e)
        {
            excel = TrialsWithForm.Excels.WriteExcel();
            Database = excel.ActiveWorkbook;
            Excel.Worksheet actsheet = Database.ActiveSheet;
        }
        //getGroupBoxes allows the groupboxes/riskmatrix to be taken as a global variable. This was not made as a public groupbox because the group boxes were made after the compile time so that it doesnt exist when
        // it first runs.
        public GroupBox[][] getGroupBoxes()
        {
            GroupBox[] textBoxesRowA = { GB1, GB2, GB3, GB4, GB5 };
            GroupBox[] textBoxesRowB = { GB6, GB7, GB8, GB9, GB10 };
            GroupBox[] textBoxesRowC = { GB11, GB12, GB13, GB14, GB15 };
            GroupBox[] textBoxesRowD = { GB16, GB17, GB18, GB19, GB20 };
            GroupBox[] textBoxesRowE = { GB211, GB22, GB23, GB24, GB25 };
            GroupBox[][] textBoxes = { textBoxesRowA, textBoxesRowB, textBoxesRowC, textBoxesRowD, textBoxesRowE };
            return textBoxes;
        }
        //Not important. Self generated code made from Windows.
        private void Area_Click(object sender, EventArgs e)
        {

        }
        //ClearContents is made so that it would clear the entire risk matrix in the event of a new risk entry being loaded into the database or when the database in memory is closed
        public void ClearContents()
        {
            GroupBox[][] textBoxes = getGroupBoxes();
            for (int i = 0; i<=4; i++)
            {
                for(int j = 0; j<=4; j ++)
                {
                    textBoxes[i][j].Controls.Clear();
                }
            }
        }
        //Point_Click allows the addon application to add new risk icons in the project by providing the necessary information to be placed into Revit such as the required Risk icon and colour
        //It also creates the neccessary GUI in the point control tab that allows the user to place the necessary information regarding the risk icon position.
        // It is important to note that the change in the inputInformation tab with the change in severity and likelihood as well as the risk descipline could change the risk icon that is used
        // This make it so that you could relate one risk entry with multiple differen risk icons giving the multidisciplinary risk feature to the addon application.
        // The pointControl tab would close after the OK button on the point control is pressed.
        public void Point_Click(object sender, EventArgs e)
        {
            int flag = 0; int Numtab = tabControl.TabCount;

            if (S1.Checked) { severity = S1.Text; }
            else if (S2.Checked) { severity = S2.Text; }
            else if (S3.Checked) { severity = S3.Text; }
            else if (S4.Checked) { severity = S4.Text; }
            else if (S5.Checked) { severity = S5.Text; }

            if (L1.Checked) { likelihood = L1.Text; }
            else if (L2.Checked) { likelihood = L2.Text; }
            else if (L3.Checked) { likelihood = L3.Text; }
            else if (L4.Checked) { likelihood = L4.Text; }
            else if (L5.Checked) { likelihood = L5.Text; }
            string colours = IdentifyRiskIcon(severity, likelihood);
            string discipline = RiskDisci.Text;
            fsFamilyName = string.Concat("_" + discipline + "_" + colours + "_");
            fsName = string.Concat("_" + discipline + "_" + colours + "_");
            for (int pos = 0; pos < Numtab; pos++)
            {
                if (tabControl.TabPages[pos].Name == "pointControl")
                { flag = 1; }
            }
            if (flag != 1)
            {
                tabControl.TabPages.Insert(2, pointControl);


                riskPointID.Location = new System.Drawing.Point(32, 218);
                riskPointID.Name = "riskPoint";
                riskPointID.Size = new System.Drawing.Size(496, 298);
                riskPointID.ReadOnly = true;
                riskPointID.Text = " ";
                riskPointID.Multiline = true;
                pointControl.Controls.Add(riskPointID);

                Button newPointer = new System.Windows.Forms.Button();
                newPointer.Location = new System.Drawing.Point(799, 106);
                newPointer.Name = "newPointer";
                newPointer.Text = "NewRiskPoint";
                newPointer.Size = new System.Drawing.Size(376, 113);
                newPointer.Click += new System.EventHandler(newPointer_Click);
                pointControl.Controls.Add(newPointer);

                Label pointLabel = new Label();
                pointLabel.Location = new System.Drawing.Point(32, 106);
                pointLabel.Size = new System.Drawing.Size(126, 25);
                pointLabel.Text = "Risk Points";
                pointControl.Controls.Add(pointLabel);

                Button point_OK = new Button();
                point_OK.Location = new System.Drawing.Point(799, 343);
                point_OK.Name = "Point_OK";
                point_OK.Size = new System.Drawing.Size(376, 113);
                point_OK.Text = "OK";
                point_OK.UseVisualStyleBackColor = true;
                point_OK.Click += new System.EventHandler(point_OK_Click);
                pointControl.Controls.Add(point_OK);

            }
            tabControl.SelectedTab = pointControl;
        }
        // Allows the ability to work from the addon application to operate with the Revit project file in the active document and thus the active window
        // newPointer_Click allows the command of new risk icons to be placed in the project file
        private void newPointer_Click(object sender, EventArgs e)
        {
            Points(commandData.Application);
        }
        // point_OK_Click relates to the OK button on the pointControl tab that sends the type of risk and the element ids back to the inputInformation tab as well as removing the pointControl tab 
        //from the main addon application to signify the successful transfer of information from Revit to the main application
        public void point_OK_Click(object sender, EventArgs e)
        {
            elementBox.Text = riskPointID.Text;
            typeBox.Text = "Point";
            elementBox.ReadOnly = true;
            typeBox.ReadOnly = true;
            tabControl.TabPages.Remove(pointControl);

        }
        // Allows the ability to work from the addon application to operate with the Revit project file in the active document and thus the active window
        // Member_click relates to the command of highlighting existing risk icons or elements that could be associated with risk within th  project
        // This could be elements that are risk icons or existing and propoer elements that is on the project such as foundation slab or beams.
        private void Member_Click(object sender, EventArgs e)
        {
            elementBox.Text = " ";
            Members(commandData.Application);
            elementBox.Text = member;
            typeBox.Text = "Member";
            member = " ";
        }
        //GetExcel allows other MS VS namespaces to access the excel application so they could make necessary changes to read and write the excel database. This is not necessary because excel is usually shown as a 
        // global variable so calling this method could be achieved by WindowsFormsApp3.TabulatedForms.excel
        public static Excel.Application GetExcel()
        {
            return excel;
        }
        // Identify the risk Icon relates the numerical value of the severity and likelihood placed by the user to the severity and liklihood index within the risk matrix in the RIM tab.
        // It is important to note that the risk and likelihood values that the user states starts at 1 and arrays in C# starts at 0
        public string IdentifyRiskIcon(string riskSeverity, string riskLikelihood)
        {
            int numSeverity = Convert.ToInt32(riskSeverity);
            int numLikelihood = Convert.ToInt32(riskLikelihood);

            GroupBox[][] boxes = getGroupBoxes();
            System.Drawing.Color backRiskColor = boxes[numSeverity - 1][numLikelihood - 1].BackColor;
            string colours = backRiskColor.Name;
            return colours;
        }

        // Result Members relates to the command in Revit that allows you to select existing elements and store the element's Id that is unique to that element. This allows the designer to manipulate which element
        // are highlighted.
        public Result Members(UIApplication uiapp)
        {
            Autodesk.Revit.UI.UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<string> listnum = new List<string>();


            IList<Reference> pickedObjs = uidoc.Selection.PickObjects(ObjectType.Element, "Select Elements");
            List<ElementId> ids = (from Reference r in pickedObjs select r.ElementId).ToList();
            using (Transaction tx = new Transaction(doc))
            {
                StringBuilder sb = new StringBuilder();
                tx.Start("transaction");
                if (pickedObjs != null && pickedObjs.Count > 0)
                {
                    foreach (ElementId eid in ids)
                    {
                        Element e = doc.GetElement(eid);
                        sb.Append("\n" + eid.IntegerValue + "\t" + e.Name);
                        listnum.Add(eid.IntegerValue + "");
                    }

                }
                tx.Commit();
                listnum.ForEach(delegate (String name)
                {
                    member = member + name + " ";
                });
            }
            return Result.Succeeded;

        }
        // RealString works in conjunction with Result Points that changes the x,y,z point that is taken as a double and convert it into a string of 2 decimal places
        public static string RealString(double a)
        {
            return a.ToString("0.##");
        }
        //PointString works in conjunction with Result Points and takes the 3D point that is coded in the system in imperical values and changes them into metric values.
        public static string PointString(XYZ p)
        {
            return string.Format("({0},{1},{2})",
              RealString(p.X * 304.8),
              RealString(p.Y * 304.8),
              RealString(p.Z * 304.8));
        }
        //PickFaceSetWorkPlaneAndPickPoint works in conjunction with Result Points that allows the user to specify one of the coordinates that are used when choosing a point in 3D space.
        // It is important to note that to specify the workplane that you owuld have to choose a particular face of an element and so you cannot choose datums or input a certain elevation
        bool PickFaceSetWorkPlaneAndPickPoint(UIDocument uidoc, out XYZ point_in_3d)
        {

            point_in_3d = null;

            Document doc = uidoc.Document;
            Reference r = uidoc.Selection.PickObject(
              ObjectType.Face,
              "Please select a planar face to define work plane");

            Element e = doc.GetElement(r.ElementId);

            {
                PlanarFace face
                  = e.GetGeometryObjectFromReference(r)
                    as PlanarFace;
                try
                {
                    point_in_3d = uidoc.Selection.PickPoint(
                      "Please pick a point on the plane"
                      + " defined by the selected face");
                }
                catch (OperationCanceledException)
                {
                }


            }
            return null != point_in_3d;
        }
        //Result Points  relates to the command in Revit that allows the user to make new risk icons in the Revit document and therefore the Revit active window. This is determined by choosing a work plane 
        //choose the other coordinates in 3dspace. When this is done then the command would cause a new window to be created whereby a risk icon would be place in the window along with the project.
        // The step mentioned above simulated the regeneration of the project in the active window.
        public Result Points(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            XYZ point_in_3d;

            if (PickFaceSetWorkPlaneAndPickPoint(
              uidoc, out point_in_3d))
            {
                string xCoord = RealString(point_in_3d.X);
                string yCoord = RealString(point_in_3d.Y);
                string zCoord = RealString(point_in_3d.Z);
                xyz = new XYZ(point_in_3d.X, point_in_3d.Y, point_in_3d.Z);

                CreateWindow(uidoc, doc, fsFamilyName, fsName, zCoord, xCoord, yCoord);

                return Result.Succeeded;
            }
            else
            {
                return Result.Failed;
            }
        }
        //CreateWindow works in conjunction with Result Points and take in the necessary information such as the XYZ coordinate and the risk icon as dicated by the severity, likelihood and risk descipline 
        // and inputs the risk icon in the specifed XYZ corodinate. This then regenerates the active window
        public void CreateWindow(UIDocument uidoc, Document doc, string fsFamilyName, string fsName, string zCoord, string xCoord, string yCoord)
        {
            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).
                 OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                         where (fs.Family.Name == fsFamilyName)
                                         select fs).First();

            using (Transaction t = new Transaction(doc, "Create window"))
            {
                ICollection<ElementId> ids = new List<ElementId>();
                t.Start();

                if (!familySymbol.IsActive)
                {

                    familySymbol.Activate();
                    doc.Regenerate();
                }
                Level elevation = Level.Create(doc, xyz.Z);
                ElementId elevationId = elevation.Id;
                ids.Add(elevationId);

                FamilyInstance window = doc.Create.NewFamilyInstance(xyz, familySymbol, elevation, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                ElementId ele = window.Id;
                riskPointID.Text = riskPointID.Text + "" + ele + " ";
                t.Commit();
            }
        }
        // Allows the ability to work from the addon application to operate with the Revit project file in the active document and thus the active window
        // newPoint_Click allows the command of new risk icons to be placed in the project file.
        // old version and outdated
        private void newPoint_Click(object sender, EventArgs e)
        {
            Points(commandData.Application);
        }
        //Actives with the activation of a clicking of a linklabel. ChangeReRiskDes gets the relevant information of the risk icon and displays it on the risk description text box located on 
        //the RIM tab of the addon applicaton
        public static void ChangeReRiskDes(int i)
        {
            var xlSheet = (Excel.Worksheet)Database.ActiveSheet;
            var selectedIds = xlSheet.Cells[i, 1].Value;
            
            List<int> selectedints = new List<int>();

            selectedints = TrialsWithTabForms.ExcelsWithTabsForms.getBetweenints(selectedIds, " ", " ");
            GetName(selectedints);
            reriskmessage = reriskmessage + "\r\n\r\n" +xlSheet.Cells[i, 5].Value;
            ReRiskDes.Text = reriskmessage;
            reriskmessage = "";
            ReRiskDes.ReadOnly = true;
        }
        // Allows the ability to work from the addon application to operate with the Revit project file in the active document and thus the active window
        // GetName gets information of the risk icon that is used such as the risk discipline and the colour of the risk icon
        public static void GetName(List<int> eltrials)
        {

            Names(WindowsFormsApp3.TabulatedForms.commandData, eltrials);
        }

        public static Result Names(ExternalCommandData eData, List<int> eletrials)
        {
            UIApplication uiapp = eData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            List<string> riskCatCol = new List<string>();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Transaction Name");
                ICollection<ElementId> ids = new List<ElementId>();
                foreach (int element in eletrials)
                {
                    ElementId id = new ElementId(element);
                    Element eleFromId = doc.GetElement(id);
                    var strings = eleFromId.Name; 
                    riskCatCol = TrialsWithTabForms.ExcelsWithTabsForms.getBetweenstrings(strings, "_", "_");
                    reriskmessage = reriskmessage + "\t\t" + "RiskType: " + riskCatCol[0] + "\t" + "RiskColour: " + riskCatCol[1];
                }

                tx.Commit();
            }
            return Result.Succeeded;
        }
        //Not important. Self generated code made from Windows.
        private void ReRisk_TextChanged(object sender, EventArgs e)
        {

        }

        private void Type_Click(object sender, EventArgs e)
        {

        }
    }
}