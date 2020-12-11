using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YlExcelImport.Enum;
using YlExcelImport.Models;

namespace YlExcelImport
{
  public  class ExcelHelper
    {

        private HSSFWorkbook workbook;
        private ISheet sheet;
        private DataTable formDt; //要求为两列，第一列为Filed，第二列值Value
        public ExcelHelper(HSSFWorkbook workbook, ISheet sheet, DataTable dt)
        {
            this.workbook = workbook;
            this.sheet = sheet;
            this.formDt = dt;
        }


        public void CreateFormHeader(FormHeader formHeader)
        {
            if (formHeader.FormFields == null || formHeader.FormFields.Count == 0)
            {
                throw new Exception("请配置json文件中的‘FormFields’项");
            }
            var row = sheet.CreateRow(formHeader.StratRow);

            //获取表单表格需要占据的行号
          var rowIndexList=  formHeader.FormFields.Select(s => s.RowIndex).Distinct().ToList();
            //创建行
            rowIndexList.ForEach(index =>
            {
                sheet.CreateRow(index);
            });

            //创建单元格
            formHeader.FormFields.ForEach(item =>
            {
                CreateFormCell(item);
            });

            //合并单元格


        }
        public void CreateFormCell(FormCell formCell)
        {
            var row = sheet.GetRow(formCell.RowIndex);
            ICell preCell;  //标题
            //创建头
            if (formCell.PreCell != null)
            {
                preCell = CreateCell(row,formCell.PreCell);
            }
            else
            {
                preCell = CreateCell(row, formCell.ColumnsIndex - 1);
                var style = CreateCellStyle(formCell.CellStyle);
                preCell.CellStyle = style;
            }
            preCell.SetCellType(CellType.String);
            preCell.SetCellValue(formCell.Name);
            //创建值部分
            formCell.ColumnsIndex = formCell.ColumnsIndex;
            var cell = CreateCell(row,formCell);
            string val = "";
            if (formCell.FixedValue != null)
            {
                val = formCell.FixedValue.ToString();
            }
            else
            {
                val = GetFiledValue(formCell.Filed);
            }
            cell.SetCellValue(val);
        }



        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="tableHeader"></param>
        public void CreateTableHeader(TableHeader tableHeader)
        {
            var row = sheet.CreateRow(tableHeader.StratRow);
            if (tableHeader.Columns == null || tableHeader.Columns.Count == 0)
            {
                throw new Exception("请配置表头列名！");
            }
            //创建表头样式
            var style = CreateHeaderStyle(tableHeader);
            IRow row2 = null;
            if (tableHeader.Columns.Any(a => a.ChildHeaders != null && a.ChildHeaders.Count > 0))
            {
                row2 = sheet.CreateRow(tableHeader.StratRow + 1);
            }
            int totalLen = 0;   //已绘制的长度
            tableHeader.Columns.OrderBy(o => o.OrderNum).ToList().ForEach(item =>
                {
                    item.ColumnsIndex = tableHeader.StratColumn + totalLen;
                    var hasChild = (item.ChildHeaders != null && item.ChildHeaders.Count > 0);
                    var cell = CreateCell(row, item.ColumnsIndex);
                    if (item.CellStyle != null)
                    {
                        style = CreateCellStyle(item.CellStyle);
                    }
                    cell.CellStyle = style;
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(item.Name);

                    if (row2 != null && !hasChild)  //表示有表头嵌套
                    {
                        var temCell = CreateCell(row2, item.ColumnsIndex);  //用于合并单元格后样式问题处理
                        temCell.CellStyle = style;
                    }

                    totalLen++;
                    if (hasChild)
                    {
                        int childIndex = 0;
                        totalLen = totalLen + (item.ChildHeaders.Count - 1);
                        //绘制子表头
                        item.ChildHeaders.ForEach(child =>
                            {
                                child.ColumnsIndex = item.ColumnsIndex + childIndex;
                                var cell1 = CreateCell(row2, child);
                                var temCell1 = CreateCell(row, child);
                                if (child.CellStyle != null)
                                {
                                    style = CreateCellStyle(child.CellStyle);
                                }
                                cell1.CellStyle = style;
                                cell1.SetCellType(CellType.String);
                                cell1.SetCellValue(child.Name);
                                temCell1.CellStyle = style;           //用于合并后的样式问题处理
                                temCell1.SetCellType(CellType.String);
                                temCell1.SetCellValue(item.Name);

                                childIndex++;
                            });
                    }
                });
            //合并单元格
            if (row2 != null)
            {
                tableHeader.Columns.OrderBy(o => o.OrderNum).ToList().ForEach(item =>
                {
                    var hasChild = (item.ChildHeaders != null && item.ChildHeaders.Count > 0);
                    if (hasChild)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(tableHeader.StratRow, tableHeader.StratRow, item.ColumnsIndex, item.ColumnsIndex + item.ChildHeaders.Count - 1));
                    }
                    else
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(tableHeader.StratRow, tableHeader.StratRow + 1, item.ColumnsIndex, item.ColumnsIndex));
                    }
                });
            }
        }







        #region 公共

        public string GetFiledValue(string field)
        {
            if (formDt == null)
            {
                return "";
            }
            if (formDt.Rows.Count == 0)
            {
                return "";
            }
            if (string.IsNullOrEmpty(field))
            {
                return "";
            }
            var dr = formDt.AsEnumerable().Where(w => w["Filed"].ToString() == field).FirstOrDefault();
            if (dr == null)
            {
                return "";
            }
            return dr[1].ToString();
        }




        /// <summary>
        /// 检查颜色是否
        /// </summary>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public bool CheckColor(short fontColor)
        {
            var name = ColorEnum.GetName(typeof(ColorEnum), fontColor);
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }
            return true;
        }



        /// <summary>
        /// 创建默认单元格样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ICellStyle CreateDefaultCellStyle()
        {
            var headStyle = workbook.CreateCellStyle();
            headStyle.Alignment = HorizontalAlignment.Center;  //水平居中
            headStyle.VerticalAlignment = VerticalAlignment.Center; //垂直居中
            return headStyle;
        }
        /// <summary>
        /// 添加边框
        /// </summary>
        /// <param name="headStyle"></param>
        public void AddAllBorder(ICellStyle headStyle)
        {
            headStyle.BorderTop = BorderStyle.Thin;  //上
            headStyle.BorderBottom = BorderStyle.Thin;//下
            headStyle.BorderLeft = BorderStyle.Thin;//左
            headStyle.BorderRight = BorderStyle.Thin;//右
        }


        /// <summary>
        /// 创建默认font
        /// </summary>
        /// <returns></returns>
        public IFont CreateDefaultFont()
        {
            var font = workbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 11;
            return font;
        }

        /// <summary>
        /// 创建单元格样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="hasBorder"></param>
        /// <returns></returns>
        public ICellStyle CreateCellStyle(CellStyle cellStyle)
        {
            var headStyle = CreateDefaultCellStyle();
            var font = CreateDefaultFont();
            if (cellStyle != null)
            {
                if (cellStyle.FontSize > 0)
                {
                    font.FontHeightInPoints = (short)cellStyle.FontSize;
                }
                if (cellStyle.IsBold)
                {
                    font.Boldweight = 700;
                }
                if (CheckColor(cellStyle.FontColor))
                {
                    font.Color = cellStyle.FontColor;
                }
                //边框
                if (cellStyle.Border == (int)BorderEnum.All)
                {
                    AddAllBorder(headStyle);
                }
                else if (cellStyle.Border == (int)BorderEnum.Cus)
                {
                    headStyle.BorderTop = cellStyle.BorderTop ? BorderStyle.Thin : BorderStyle.None;  //上
                    headStyle.BorderBottom = cellStyle.BorderBottom ? BorderStyle.Thin : BorderStyle.None;//下
                    headStyle.BorderLeft = cellStyle.BorderLeft ? BorderStyle.Thin : BorderStyle.None;//左
                    headStyle.BorderRight = cellStyle.BorderRight ? BorderStyle.Thin : BorderStyle.None;//右
                }
                //背景色
                if (CheckColor(cellStyle.BackGroundColor))
                {
                    headStyle.FillBackgroundColor = cellStyle.BackGroundColor;
                    headStyle.FillForegroundColor = cellStyle.BackGroundColor;
                    headStyle.FillPattern = FillPattern.SolidForeground;
                }
            }
            headStyle.SetFont(font);
            return headStyle;
        }

        /// <summary>
        /// 创建表头通用样式
        /// </summary>
        /// <param name="tableHeader"></param>
        /// <returns></returns>
        public ICellStyle CreateHeaderStyle(TableHeader tableHeader)
        {
           var style = CreateDefaultCellStyle();
            //背景色
            if (tableHeader.BackGroundColor>0&& CheckColor(tableHeader.BackGroundColor))
            {
                style.FillBackgroundColor = tableHeader.BackGroundColor;
                style.FillForegroundColor = tableHeader.BackGroundColor;
                style.FillPattern = FillPattern.SolidForeground;
            }
            //边框
            if (tableHeader.HasBorder)
            {
                AddAllBorder(style);
            }
            var font = CreateDefaultFont();
            //字体颜色
            if (tableHeader.FontColor>0 && CheckColor(tableHeader.FontColor))
            {
                font.Color = tableHeader.FontColor;
            }
            //字体大小
            if (tableHeader.FontSize > 0)
            {
                font.FontHeightInPoints =tableHeader.FontSize;
            }
            //是否加粗
            if (tableHeader.IsBold)
            {
                font.Boldweight = 700;
            }
            style.SetFont(font);
            return style;
        }


        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="baseCell"></param>
        /// <returns></returns>
        public ICell CreateCell(IRow row, BaseCell baseCell)
        {
            var cellIndex = 0;
            if (baseCell != null)
            {
                cellIndex = baseCell.ColumnsIndex;
            }
            var cell = row.CreateCell(cellIndex);
            return cell;
        }
        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cIndex"></param>
        /// <returns></returns>
        public ICell CreateCell(IRow row, int cIndex)
        {

            return row.CreateCell(cIndex);
        }

        #endregion









    }
}
