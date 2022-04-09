using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreatePlugin_4
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            Level level1 = GetLevel(doc);

            Transaction transaction = new Transaction(doc, "Create");
            transaction.Start();

            List<Wall> walls = CreateWall(doc, level1.Id, 10, 5, 3);
            AddDoor(doc, level1, walls[0]);
            AddWindows(doc, level1, walls);

            transaction.Commit();


            return Result.Succeeded;
        }

        private void AddWindows(Document doc, Level level1, List<Wall> walls)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0610 x 1220 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            for (int i = 0; i < walls.Count(); i++)
            {
                LocationCurve hostCurve = walls[i].Location as LocationCurve;
                XYZ point1 = hostCurve.Curve.GetEndPoint(0);
                XYZ point2 = hostCurve.Curve.GetEndPoint(1);
                XYZ z = new XYZ(0, 0, UnitUtils.ConvertToInternalUnits(0.9, UnitTypeId.Meters));
                XYZ point = (point1 + point2) / 2 + z;

                if (!windowType.IsActive)
                    windowType.Activate();

                IList<ElementId> inserts = walls[i].FindInserts(true, true, true, true);
                if (0 == inserts.Count)
                {
                    doc.Create.NewFamilyInstance(point, windowType, walls[i], level1, StructuralType.NonStructural);
                }
            }
        }

        private void AddDoor(Document doc, Level level1, Wall wall)
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

            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(point, doorType, wall, level1, StructuralType.NonStructural);
        }

        public Level GetLevel(Document doc)
        {
            Level level1 = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList()
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();

            return level1;
        }

        public List<Wall> CreateWall(Document doc, ElementId levelId, double width, double depth, double height)
        {
            double wd = UnitUtils.ConvertToInternalUnits(width, UnitTypeId.Meters);
            double dp = UnitUtils.ConvertToInternalUnits(depth, UnitTypeId.Meters);
            double hg = UnitUtils.ConvertToInternalUnits(height, UnitTypeId.Meters);

            double dx = wd / 2;
            double dy = dp / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            WallType wall1 = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .OfType<WallType>()
                .ToList()
                .Where(x => x.Name.Equals("Наружный - Стена из кирпича с наружным слоем лицевого кирпича толщиной 380 мм и Кирпич фасадный - 250ммx65мм оштукатуренная 25 мм"))
                .FirstOrDefault();

            List<Wall> walls = new List<Wall>();

            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, wall1.Id, levelId, hg, 0, true, false);
                walls.Add(wall);
            }

            return walls;
        }
    }
}
