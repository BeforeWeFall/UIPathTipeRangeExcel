using ClosedXML.Excel;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace My.Activities.TypeCell
{
    public class CellType : CodeActivity
    {

        public enum typeCell
        {
            General = 0,
            Number=1,
            NumberWithCom=2,
            Money =5,
            MoneyWithCom = 7,
            Percent=9,
            PercentWithCom=10,
            dd_MM_yyyy=14,
            dd_MMM_yy=15,
            dd_MMM=16,
            MMM_yy=17,
            hh_mm_AM=18,
            hh_mm_ss_AM=19,
            h_mm=20,
            h_mm_ss=21,
            dd_MM_yyyy_h_mm=22
        }

        private IXLWorkbook book;
        private IXLWorksheet worksheet;

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> PathExcel { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> SheetName { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Cell { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public typeCell Type { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string path = Regex.Replace(PathExcel.Get(context), @"[^\P{C}\n]+","");
            string sheetName = SheetName.Get(context);
            string cell = Cell.Get(context);
            SetTip(path, sheetName, cell, (int)Type);
        }

        public void SetTip(string path, string sheetName, string cell, int format )
        {
            if (File.Exists(path))
            {
                book = new XLWorkbook(path);

                try
                {
                    worksheet = book.Worksheet(sheetName);
                }
                catch
                {
                    worksheet = book.AddWorksheet(sheetName);
                }
            }
            else
            {
                book = new XLWorkbook();
                worksheet = book.AddWorksheet(sheetName);
            }

            if (!cell.Contains(":"))
                TipCell(cell, format);
            else
                TipRange(cell, format);
            book.Save();
            //wb.SaveAs(filePath);
        }

        private void TipCell(string target, int format)
        {
            worksheet.Cell(target).Style.NumberFormat.NumberFormatId = format; 
        }
        private void TipRange(string target, int format)
        {
            string[] range = target.Split(':');

            IXLRange rangeXL;
            if (string.IsNullOrWhiteSpace(range[1]))
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), GetAlfb(worksheet.RangeUsed().FirstRowUsed().CellCount() - 1) + (worksheet.RangeUsed().RowCount()));
            }
            else
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), range[1].ToUpper());
            }
            rangeXL.Style.NumberFormat.NumberFormatId = format;
        }
        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)(num % 90)).ToString() : ((char)(065 + num)).ToString();
        }
    }
}
