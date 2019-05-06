using System;
using System.Windows.Media.Imaging;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Lighting;

namespace Luminotecnico
{
    public class CsAddPanel : IExternalApplication
    {
        // Both OnStartup and OnShutdown must be implemented as public method
        public Result OnStartup(UIControlledApplication application)
        {
            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Cálculo Luminotécnico");

            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("cmdHelloWorld",
               "Hello World", thisAssemblyPath, "Luminotecnico.Main");

            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

            // Optionally, other properties may be assigned to the button
            // a) tool-tip
            pushButton.ToolTip = "Verifica se o espaço está de acordo com a norma ABNT15.575";

            // b) large bitmap
            Uri uriImage = new Uri(@"C:\Users\emill\source\repos\luminotecnico\luminotecnico\lamp.png");
         BitmapImage largeImage = new BitmapImage(uriImage);
         pushButton.LargeImage = largeImage;

         return Result.Succeeded;
      }

      public Result OnShutdown(UIControlledApplication application)
      {
         // nothing to clean up in this simple case
         return Result.Succeeded;
      }
   }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            TaskDialog.Show("Revit","Para realizar os cálculos é necessário selecionar um ponto em um espaço definido anteriormente.");
            Selection sel = uiapp.ActiveUIDocument.Selection;
            XYZ eye = sel.PickPoint("Por favor, escolha um ponto para realizar os cálculos Lumininotécnicos");
            Space space = doc.GetSpaceAtPoint(eye);
            if (space != null)
            {
                Forms.Luminotecnico luminotecnico = new Forms.Luminotecnico(space);
                luminotecnico.ShowDialog();
            }
            else
                TaskDialog.Show("Revit", "Selecione um ponto no espaço criado.");

            return Result.Succeeded;
        }


    }
}
