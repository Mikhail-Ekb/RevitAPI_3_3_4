using Aspose.Cells.Charts;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_3_3_4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Код создания параметра
            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));
            using (Transaction ts = new Transaction(doc, "Задайте параметр"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Внешний/Внутренний диаметр", categorySet, BuiltInParameterGroup.PG_DATA, true);

                ts.Commit();
            }

            //Код заполнения нового параметра
            IList<Pipe> pipeList = null;
            pipeList = new FilteredElementCollector(doc, doc.ActiveView.Id)
            .OfClass(typeof(Pipe))
            .Cast<Pipe>()
            .ToList();

            using (Transaction ts = new Transaction(doc, "Задайте параметр"))
            {
                ts.Start();
                foreach (var pipeInstance in pipeList)
                {                    
                    if (pipeInstance is Pipe)
                    {
                        Parameter rMaxParametr = pipeInstance.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                        Parameter rMinParametr = pipeInstance.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                        
                        if (rMaxParametr.StorageType == StorageType.Double || rMinParametr.StorageType == StorageType.Double)
                        {
                            double rMaxhValue = UnitUtils.ConvertFromInternalUnits(rMaxParametr.AsDouble(), /*UnitTypeId.Meters*/ DisplayUnitType.DUT_METERS);
                            double rMax = rMaxhValue/*.AsDouble()*/;                                                    
                        
                            double rMinhValue = UnitUtils.ConvertFromInternalUnits(rMaxParametr.AsDouble(), /*UnitTypeId.Meters*/ DisplayUnitType.DUT_METERS);
                            double rMin = rMinhValue/*.AsDouble()*/;
                            Parameter rMaxMinParametr = pipeInstance.LookupParameter("Внешний/Внутренний диаметр");
                            rMaxMinParametr.Set($"Труба {rMax}/{rMin}");
                        }
                        
                    }
                }

                ts.Commit();
            }

            return Result.Succeeded;
        }

        // Метод для кода создания параметра
        private void CreateSharedParameter(Application application, Document doc,
            string parameterName, CategorySet categorySet,
            BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile defFile = application.OpenSharedParameterFile();

            if (defFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }
            Definition definition = defFile.Groups
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
            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }

    }
}