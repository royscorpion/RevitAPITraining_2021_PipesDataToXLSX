using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_2021_PipesDataToXLSX
{
    public static class SheetExts
    {
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)
        {
            var cellReference = new CellReference(rowIndex, columnIndex); //создаем переменную ссылки на ячейку
            var row = sheet.GetRow(cellReference.Row); //создаем переменную строки
            if (row == null) //если строки нет
                row = sheet.CreateRow(cellReference.Row); //создаем новую строку

            var cell = row.GetCell(cellReference.Col); //создаем ссылку на конкретную ячейку
            if (cell == null) //если ячейка пустая
                cell = row.CreateCell(cellReference.Col); //то создаем ее

            if(value is string)
            {
                cell.SetCellValue((string)(object)value);
            }
            else if (value is double)
            {
                cell.SetCellValue((double)(object)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)(object)value);
            }

        }
    }
}
