using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            Document vkDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ВК")).FirstOrDefault();
            if (ovDoc == null || vkDoc == null)
            {
                if (ovDoc == null)
                {
                    TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                    return Result.Cancelled;
                }
                else if (vkDoc == null)
                {
                    TaskDialog.Show("Ошибка", "Не найден ВК файл");
                    return Result.Cancelled;
                }
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие")|| x.FamilyName.Equals("Отверстие в полу"))
                .FirstOrDefault();
            FamilySymbol familySymbolFloor = new FilteredElementCollector(arDoc)
               .OfClass(typeof(FamilySymbol))
               .OfCategory(BuiltInCategory.OST_GenericModel)
               .OfType<FamilySymbol>()
               .Where(x => x.FamilyName.Equals("Отверстие в полу"))
               .FirstOrDefault();

            if (familySymbol == null|| familySymbolFloor==null)
            {
                TaskDialog.Show("Ошибка", "Семейство \"Отверстие\" не найдено");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(vkDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);


            Transaction ts = new Transaction(arDoc);
            ts.Start("Создать отверстие в стене");

            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();

            }
            DuctHole(arDoc, ducts, referenceIntersector, familySymbol);
            PipeHole(arDoc, pipes, referenceIntersector, familySymbol);
           

            ts.Commit();

            ReferenceIntersector referenceIntersector1 = new ReferenceIntersector(new ElementClassFilter(typeof(Floor)), FindReferenceTarget.Element, view3D);
            Transaction ts1 = new Transaction(arDoc);
            ts1.Start("Создать отверстие в перекрытии");

            if (!familySymbolFloor.IsActive)
            {
                familySymbolFloor.Activate();
            }
            
            PipeHoleFioor(arDoc, pipes, referenceIntersector1, familySymbolFloor);

            ts1.Commit();
            return Result.Succeeded;

        }
       
        public void PipeHoleFioor(Document arDoc, List<Pipe> pipes, ReferenceIntersector referenceIntersector1, FamilySymbol familySymbolFloor)
        {
            foreach (Pipe p in pipes)
            {
                Line line = (p.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector1.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                

                foreach (ReferenceWithContext inter in intersections)
                {
                    double proximity = inter.Proximity;
                    Reference reference = inter.GetReference();


                    XYZ pointHole = point + (direction * proximity);
                    Floor floor = arDoc.GetElement(reference.ElementId) as Floor;
                  //  Level level1 = arDoc.GetElement(floor.LevelId) as Level;

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbolFloor, floor, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter heigth = hole.LookupParameter("Высота");
                    width.Set(p.Diameter);
                    heigth.Set(p.Diameter);

                }

            }
        }

        public void PipeHole(Document arDoc, List<Pipe> pipes, ReferenceIntersector referenceIntersector, FamilySymbol familySymbol)
        {
            foreach (Pipe p in pipes)
            {
                Line line = (p.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();


                foreach (ReferenceWithContext inter in intersections)
                {
                    double proximity = inter.Proximity;
                    Reference reference = inter.GetReference();

                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;

                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter heigth = hole.LookupParameter("Высота");
                    width.Set(p.Diameter);
                    heigth.Set(p.Diameter);

                }

            }
        }

        public void DuctHole(Document arDoc, List<Duct> ducts, ReferenceIntersector referenceIntersector, FamilySymbol familySymbol)
        {
            foreach (Duct d in ducts)
            {
                Line line = (d.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();


                foreach (ReferenceWithContext inter in intersections)
                {
                    double proximity = inter.Proximity;
                    Reference reference = inter.GetReference();

                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter heigth = hole.LookupParameter("Высота");
                    width.Set(d.Diameter);
                    heigth.Set(d.Diameter);
                }
            }
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
