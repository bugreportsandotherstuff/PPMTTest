using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPMTTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var newFile = $"PPMTtest_{DateTime.Now:yyyy.MM.dd_HH.mm.ss.fff}.xlsx";

            using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
            {

                IWorkbook workbook = new XSSFWorkbook();

                ISheet sheet1 = workbook.CreateSheet("Sheet1");
               
                //PPMT with 4 parameters (without Fv)
                var row1 = sheet1.CreateRow(1);
                var cell11 = row1.CreateCell(1);
                cell11.SetCellValue("PPMT Without Fv parameter");
                var cell12 = row1.CreateCell(2);
                cell12.SetCellFormula("PPMT(0.002892562,1,60,25650)");

                //PPMT with 4 parameters (with Fv)
                var row2 = sheet1.CreateRow(2);
                var cell21 = row2.CreateCell(1);
                cell21.SetCellValue("PPMT With Fv parameter");
                var cell22 = row2.CreateCell(2);
                cell22.SetCellFormula("PPMT(0.002892562,1,60,25650,-7125)");


                //Debug test
                var debugValue = NPOI.SS.Formula.Functions.Finance.PPMT(0.002892562, 1, 60, 25650, -7125);
                var row3 = sheet1.CreateRow(3);
                var cell31 = row3.CreateCell(1);
                cell31.SetCellValue("Debug PPMT With Fv parameter calculated from code");
                var cell32 = row3.CreateCell(2);
                cell32.SetCellValue(debugValue);

                XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
                workbook.Write(fs);
            }

        }
    }
}
