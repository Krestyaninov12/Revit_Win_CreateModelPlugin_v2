using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateModelPlugin_v2
{
    [TransactionAttribute(TransactionMode.Manual)]

    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level1 = SelectLevel(doc, "Уровень 1");
            Level level2 = SelectLevel(doc, "Уровень 2");

            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters); // длина
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);  // глубина

            double dx = width / 2;
            double dy = depth / 2;

            Transaction ts = new Transaction(doc, "Построение модели");
            ts.Start();
            {
                List<Wall> walls = CreateWall(dx, dy, level1, level2, doc);  // построение стен и составление списка стен
                AddDoor(doc, level1, walls[1]);           //построение двери на Level1 в стене walls[1]
                AddWindow(doc, level1, walls[0]);           // построение окна на Level1 в стене walls[0]
                AddWindow(doc, level1, walls[2]);          // построение окна на Level1 в стене walls[2]
                AddWindow(doc, level1, walls[3]);	        // построение окна на Level1 в стене walls[3]
            }
            ts.Commit();

            return Result.Succeeded;
        }

        private void AddWindow(Document doc, Level level1, Wall wall)                // метод построения окна
        {
            FamilySymbol windowType = new FilteredElementCollector(doc) // через фильтр получаем тип окна (загруженного в проекте)
                 .OfClass(typeof(FamilySymbol))
                 .OfCategory(BuiltInCategory.OST_Windows)
                 .OfType<FamilySymbol>()
                 .Where(x => x.Name == "0915 x 1830 мм")
                 .Where(x => x.Family.Name == "Фиксированные")
                 .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;  // получение curve стены
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);		// левая точка curve
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);		// правая точка curve
            XYZ point = (point1 + point2) / 2;				// середина curve

            if (!windowType.IsActive)				// активация типа
            {
                windowType.Activate();
            }

            FamilyInstance window = doc.Create.NewFamilyInstance(point, windowType, wall, level1, StructuralType.NonStructural);    // устанавка окна в точку point, тип windowType, в стену  wall, на уровне level1

            window.flipFacing();  		// поворот окна

            double height = UnitUtils.ConvertToInternalUnits(1000, UnitTypeId.Millimeters);
            window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(height);  // установка высоты подоконника
        }

        private void AddDoor(Document doc, Level level1, Wall wall)         // метод построения двери
        {
            FamilySymbol doorType = new FilteredElementCollector(doc) // через фильтр получаем тип двери (загруженной в проекте)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name == "0915 x 2134 мм")
                 .Where(x => x.Family.Name == "Одиночные-Щитовые")
                 .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;   // получение curve стены
            XYZ point1 = hostCurve.Curve.GetEndPoint(0); 		        // левая точка curve
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);		        // правая точка curve
            XYZ point = (point1 + point2) / 2;				            // середина curve

            if (!doorType.IsActive)				                        // активация типа
            {
                doorType.Activate();
            }

            doc.Create.NewFamilyInstance(point, doorType, wall, level1, StructuralType.NonStructural); // установка двери в точку point, тип doorType, в стену  wall, на уровне level1

        }

        public Level SelectLevel(Document doc, string levelName)  // метод отбора уровня
        {
            List<Level> listLevels = new FilteredElementCollector(doc)  // фильтр по уровням
                        .OfClass(typeof(Level))
                         .OfType<Level>()
                          .ToList();
            Level SelectLevel = listLevels
                            .Where(x => x.Name.Equals(levelName))       // фильтр по имени уровня
                            .OfType<Level>()
                            .FirstOrDefault();
            return SelectLevel;
        }

        public List<Wall> CreateWall(double dx, double dy, Level lowLevel, Level highLevel, Document doc) // метод создания стен
        {
            List<Wall> walls = new List<Wall>();                    // список стен
            List<XYZ> points = new List<XYZ>();                     // список точек

            for (int i = 0; i < 4; i++)
            {
                points.Add(new XYZ(-dx, -dy, 0));
                points.Add(new XYZ(dx, -dy, 0));
                points.Add(new XYZ(dx, dy, 0));
                points.Add(new XYZ(-dx, dy, 0));
                points.Add(new XYZ(-dx, -dy, 0));
                Line line = Line.CreateBound(points[i], points[i + 1]);   // построение линии
                Wall wall = Wall.Create(doc, line, lowLevel.Id, false);   // построение стены по линии
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(highLevel.Id);   // ограничение стены сверху
                walls.Add(wall);                                        // добавление в список созданной стены
            }
            return walls;
        }

    }





}