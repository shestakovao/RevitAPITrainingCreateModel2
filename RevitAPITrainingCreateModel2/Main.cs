using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingCreateModel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Wall> walls = CreateWalls(doc, 10000, 5000, GetLevel(doc, "Уровень 1"), GetLevel(doc, "Уровень 2"));
            AddDoor(doc, GetLevel(doc, "Уровень 1"), walls[0]);
            for (int i = 1; i < 4; i++)
            { 
                AddWindow(doc, GetLevel(doc, "Уровень 1"), walls[i], 1000); 
            }
            return Result.Succeeded;
        }

        private void AddDoor(Document doc, Level level, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;
       
            Transaction transaction = new Transaction(doc, "Добавление двери");
            transaction.Start();
            if (!doorType.IsActive) doorType.Activate();
            doc.Create.NewFamilyInstance(point, doorType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            transaction.Commit();
        }

        private void AddWindow(Document doc, Level level, Wall wall, double height)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            Transaction transaction = new Transaction(doc, "Добавление окна");
            transaction.Start();
            if (!windowType.IsActive) windowType.Activate();
            FamilyInstance wind = doc.Create.NewFamilyInstance(point, windowType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            //высота нижнего бруса
            wind.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(UnitUtils.ConvertToInternalUnits(height, UnitTypeId.Millimeters));
            transaction.Commit();
        }


        public Level GetLevel(Document doc, string inputString)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
                                    .OfClass(typeof(Level))
                                    .OfType<Level>()
                                    .ToList();

            Level level = listLevel
               .Where(x => x.Name.Equals(inputString))
               .FirstOrDefault();

            return level;
        }

        public List<Wall> CreateWalls(Document doc, double widthInput, double depthInput, Level level1, Level level2)
        {

            double width = UnitUtils.ConvertToInternalUnits(widthInput, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(depthInput, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(doc, "Построение стен");
            transaction.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }
            transaction.Commit();

            return walls;
        }

    }
}
