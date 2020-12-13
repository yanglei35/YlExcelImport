using Newtonsoft.Json;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YlExcelImport;

namespace ConsoleApp1.Test
{
   public  class ExcelTest :ExcelBase
    {
        public ExcelTest(string path) : base(path)
        {

        }

        public override void SetRowData(IRow row, ICell cell, string field, string value)
        {
            if (value.Contains('3'))
            {
                ICellStyle tem = CreateDefaultCellStyle();
                IFont font = CreateDefaultFont();
                font.Color = 10;
                font.FontHeightInPoints = 11;
                font.Boldweight = 700;
                tem.SetFont(font);
                tem.FillForegroundColor = 40;
                tem.FillPattern= FillPattern.SolidForeground;
                AddAllBorder(tem);
                cell.CellStyle = tem;
                row.Height = 40 * 20;
            }
        }
    }
}
