using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Transaction = Autodesk.Revit.DB.Transaction;

namespace RevitAPITraining_StockLength
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));

            using (Transaction ts = new Transaction(doc, "Add parameter"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Длина с запасом", categorySet, BuiltInParameterGroup.PG_DATA, true);
                ts.Commit();
            }

            double ratio = 1.1;
            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            foreach (var elem in pipes)
            {
                Parameter length = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                Parameter stockLength = elem.LookupParameter("Длина с запасом");

                double newLength = length.AsDouble() * ratio;
                double lengthInMeters = UnitUtils.ConvertFromInternalUnits(newLength, UnitTypeId.Meters);
                string result = lengthInMeters.ToString() + "м";

                using (Transaction ts = new Transaction(doc, "Set parameter"))
                {
                    ts.Start();
                    stockLength.Set(result);
                    ts.Commit();
                }
            }
            return Result.Succeeded;
        }

        private void CreateSharedParameter(
            Application application,
            Document doc,
            string parameterName,
            CategorySet categorySet,
            BuiltInParameterGroup builtInParameterGroup,
            bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = definitionFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap bindingMap = doc.ParameterBindings;
            bindingMap.Insert(definition, binding, builtInParameterGroup);
        }
    }
}
