using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_2021_PipesDataToXLSX
{
    [Transaction(TransactionMode.Manual)] //атрибут, указывающий что будет делать данное приложение - считывать данные (ReadOnly) или изменять их (Manual)
    public class Main : IExternalCommand  //переименование класса Class1 в Main (Ctrl+RR). Чтобы реализовать приложение для Revit необходимо реализовать интерфейс IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application; //обращение к приложению Revit
            UIDocument uidoc = uiapp.ActiveUIDocument; //обращение к интерфейсу текущего документа
            Document doc = uidoc.Document; //обращение к документу через интерфейс документа


            string pipeInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .Where(x => UnitUtils.ConvertFromInternalUnits(x.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Meters) >= 0.1) //фильтр по длине труб - не менее 0,1м
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "PipesData.xlsx"); //путь сохранения файла

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write)) //создаем поток по указанному пути с созданием файла и уровнем доступа на запись
            {
                IWorkbook workbook = new XSSFWorkbook(); //создаем новую книгу (файл)
                ISheet sheet = workbook.CreateSheet("Pipes"); //создаем лист

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.PipeType.Name.ToString());
                    sheet.SetCellValue(rowIndex, columnIndex: 1, UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(), UnitTypeId.Millimeters));
                    sheet.SetCellValue(rowIndex, columnIndex: 2, UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble(), UnitTypeId.Millimeters));
                    sheet.SetCellValue(rowIndex, columnIndex: 3, Math.Round(UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Meters), 2));
                    rowIndex++;
                }
                workbook.Write(stream); //сохраняем книгу
                workbook.Close();//закрываем книгу
            }

            System.Diagnostics.Process.Start(excelPath);
            return Result.Succeeded;
        }
    }
}
