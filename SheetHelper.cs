using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNPOI
{
    public static class SheetHelper
    {
        public static ICell GetCell(
            this ISheet sheet, int linha, int coluna)
        {  

            IRow row;
           // int indiceLinha = linha - 1;
            row = sheet.GetRow(linha);
            if (row == null)
                row = sheet.CreateRow(linha);

            ICell cell;
           // int indiceColuna = coluna - 1;
            cell = row.GetCell(coluna);
            if (cell == null)
                cell = row.CreateCell(coluna);

            return cell;
        }
    }
}