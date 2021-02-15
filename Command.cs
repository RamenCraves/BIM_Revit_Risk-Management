#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Text;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Media.Imaging;
using applications = System.Windows.Forms.Application;
#endregion

/// <summary>
/// TrialsWithTabForms is the main approach the user interacts with the Revit application to create the addon application instance and therefore the whole project
/// Author: Millar Sue msue619
/// 2018 p4p #122: Risk information management in Building Information Modelling (BIM)
/// </summary>
namespace TrialsWithTabForms
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class HelloPanel : IExternalApplication
    {

        // Both OnStartup and OnShutdown must be implemented as public method
        public Result OnStartup(UIControlledApplication application)
        {
            //Creates a push button to trigger a command add it to the ribbon panel.
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("#122 UoA P4P");
            string thisAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("cmdCommand",
               "TrialWithTabsForms", thisAssemblyPath, "TrialsWithTabForms.Command");

            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

            //description and the picture that is displayed on the Revit Addon tab
            pushButton.ToolTip = "UoA Part 4 Project #122";
            Uri uriImage = new Uri(@"C:\Users\mini_\Downloads\favicon.ico");
            BitmapImage largeImage = new BitmapImage(uriImage);
            pushButton.LargeImage = largeImage;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            TrialsWithForm.Excels.CloseDatabase(TrialsWithForm.Excels.excel);
            return Result.Succeeded;
        }
    }
    //Command allows TrialsWithTabForms to be generated from the pushbutton in the Revit Addon tab.
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
              ExternalCommandData commandData,
              ref string message,
              ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            WindowsFormsApp3.TabulatedForms.CreateTabWithForms(commandData);
            return Result.Succeeded;
        } 
    }
}
