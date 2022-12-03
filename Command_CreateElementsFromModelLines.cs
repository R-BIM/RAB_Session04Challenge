#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

#endregion

namespace RAB_Session04Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Command_CreateElementsFromModelLines : IExternalCommand
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

            //Create and start the t
            Transaction t = new Transaction(doc);
            t.Start("Create  model elements from model curve lines");

            //Declare variables
            WallType gen8WT = GetWallTypeByName(doc, "Generic - 8\"");
            WallType storFrrontWT = GetWallTypeByName(doc, "Storefront");
            Level myLevel = GetLevelByName(doc, "Level 1");
            MEPSystemType pipeSystemType = GetMepSystemTypeByName(doc, "Hydronic Supply");
            MEPSystemType ductSystemType = GetMepSystemTypeByName(doc, "Supply Air");
            DuctType ductType = GetDuctTypeByName(doc, "Default");
            PipeType pipeType = GetPipetTypeByName(doc, "Default");

            //Prompt the user to select elements
            IList<Element> mySelection = uidoc.Selection.PickElementsByRectangle("Please, select elements by rectangle in the screen");

            //Filter the elements for model curves
            //Loop through filtered elements (mySelection list) and based on the line's line style create elements
            foreach (Element element in mySelection)
            {
                try
                {
                    if (element is CurveElement)
                    {
                        //Cast elements as curveElements
                        CurveElement curve = element as CurveElement;

                        //Get the geometry curve of the curve element
                        Curve geomCurve = curve.GeometryCurve;

                        //Get the 3D points from geometry curve at the start of the curve
                        XYZ startPoint = curve.GeometryCurve.GetEndPoint(0);

                        //Get the 3D points from geometry curve at the end of the curve
                        XYZ endPoint = curve.GeometryCurve.GetEndPoint(1);

                        //Loop through the curves of the selection and If the line's line style name is "A-WALL" create "Generic 8" wall" wall
                        if (curve.CurveElementType == CurveElementType.ModelCurve && curve.LineStyle.Name == "A-WALL")
                        {
                            Wall wall = Wall.Create(doc, geomCurve, gen8WT.Id, myLevel.Id, 20, 0, false, false);
                            wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(10);
                        }
                        //Loop through the curves of the selection and If the line's line style name is "A-GLAZ" create "Storefront" curtain wall
                        else if (curve.CurveElementType == CurveElementType.ModelCurve && curve.LineStyle.Name == "A-GLAZ")
                        {
                            Wall wall = Wall.Create(doc, geomCurve, storFrrontWT.Id, myLevel.Id, 20, 0, false, false);
                            wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(10);
                        }
                        //Loop through the curves of the selection and If the line's line style name is "M-DUCT" create "default" rectangular duct
                        else if (curve.CurveElementType == CurveElementType.ModelCurve && curve.LineStyle.Name == "M-DUCT")
                        {
                            Duct myDuct = Duct.Create(doc, ductSystemType.Id, ductType.Id, myLevel.Id, startPoint, endPoint);
                                 myDuct.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(8.5);
                        }
                        //Loop through the curves of the selection and If the line's line style name is "P-PIPE" create "default" pipe
                        else if (curve.CurveElementType == CurveElementType.ModelCurve && curve.LineStyle.Name == "P-PIPE")
                        {
                            Pipe myPipe = Pipe.Create(doc, pipeSystemType.Id, pipeType.Id, myLevel.Id, startPoint, endPoint);
                            myPipe.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).Set(8.5);
                        }
                    }
                }
                catch (Exception e)
                {
                    message = e.Message;
                }
            }

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
        //Create a private "GetPipeTypeByName" method
        private PipeType GetPipetTypeByName(Document doc, string pipeType)
        {
            FilteredElementCollector pipeTypeCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(PipeType));
            foreach (PipeType pIPEType in pipeTypeCollector)
            {
                if (pIPEType.Name == pipeType)
                {
                    return pIPEType;
                }
            }
            return null;
        }
        //Create a private "GetDuctTypeByName" method
        private DuctType GetDuctTypeByName(Document doc, string ductType)
        {
            FilteredElementCollector ductTypeCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(DuctType));
            foreach (DuctType dUCTType in ductTypeCollector)
            {
                if (dUCTType.Name == ductType)
                {
                    return dUCTType;
                }
            }
            return null;
        }
        //Create a private "GetMepStystemTypeByName" method
        private MEPSystemType GetMepSystemTypeByName(Document doc, string mepSystemType)
        {
            FilteredElementCollector mEPSystCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystemType));
            foreach (MEPSystemType mEPSystemType in mEPSystCollector)
            {
                if (mEPSystemType.Name == mepSystemType)
                {
                    return mEPSystemType;
                }
            }
            return null;
        }
        //Create a private "GetWallTypeByName" method
        private WallType GetWallTypeByName(Document doc, string wallType)
        {

            FilteredElementCollector wtCollector = new FilteredElementCollector(doc)
            //.OfClass(typeof(WallType));
            .OfCategory(BuiltInCategory.OST_Walls).WhereElementIsElementType();

            foreach (WallType myWallType in wtCollector)
            {
                if (myWallType.Name == wallType)

                    return myWallType;
            }
            return null;
        }
        //Create a private "GetLevelByName" method
        private Level GetLevelByName(Document doc, string level)
        {
            FilteredElementCollector lvlcollector = new FilteredElementCollector(doc)
          //.OfClass(typeof(Level));
          .OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType();

            foreach (Level lvl in lvlcollector)
            {
                if (lvl.Name == level)
                    return lvl;
            }
            return null;
        }
    }
}
