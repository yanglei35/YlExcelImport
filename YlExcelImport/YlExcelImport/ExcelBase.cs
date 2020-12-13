using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YlExcelImport.Enum;
using YlExcelImport.Models;

namespace YlExcelImport
{
  public  class ExcelBase
    {
        private HSSFWorkbook workbook;
        private ISheet sheet;
        private SortedList<string,int> columConfig=new SortedList<string, int>();
        private ExcelConfig _excelConfig = null;
        private List<FormOption> _formOption = new List<FormOption>();

        public ExcelBase(string path)
        {
            this.workbook = new HSSFWorkbook();
            this.sheet = this.workbook.CreateSheet();
            ReadExcelConfig(path);
        }

        #region 暴露出的方法

        /// <summary>
        /// 普通表格导出
        /// </summary>
        /// <param name="formData"></param>
        /// <returns></returns>
        public HSSFWorkbook ExcelCommonTableExprot(DataTable formData)
        {
            if (formData == null)
            {
                throw new Exception("数据源不能为空！");
            }
            return CreateCommonTable(formData);
        }
        /// <summary>
        /// 普通表格导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public HSSFWorkbook ExcelCommonTableExprot<T>(List<T> dataList)
        {
            if (dataList == null)
            {
                throw new Exception("数据源不能为空！");
            }
            return CreateCommonTable(dataList);
        }

        /// <summary>
        /// 表单表格导出
        /// </summary>
        /// <param name="formOptions"></param>
        /// <param name="formData"></param>
        /// <returns></returns>
        public HSSFWorkbook ExcelFormTableExprot(List<FormOption> formOptions, DataTable formData)
        {
            if (formOptions == null)
            {
                throw new ArgumentNullException(nameof(formOptions));
            }

            if (formData == null)
            {
                throw new Exception("配置为表单表格时，表单的数据源不能为空！");
            }
            ExcelConfigCheck();
            return CreateFormTable(formOptions, formData);
        }

        /// <summary>
        /// 表单表格导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="formOptions"></param>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public HSSFWorkbook ExcelFormTableExprot<T>(List<FormOption> formOptions, List<T> dataList)
        {
            if (formOptions == null)
            {
                throw new ArgumentNullException(nameof(formOptions));
            }

            if (dataList == null)
            {
                throw new Exception("配置为表单表格时，表单的数据源不能为空！");
            }

            ExcelConfigCheck();
            return CreateFormTable(formOptions, dataList);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public HSSFWorkbook GetWorkBook()
        {
            if (this.workbook != null)
            {
                return this.workbook;
            }
            return null;
        }

       

        #endregion


        /// <summary>
        /// 获取excel配置对象
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private void ReadExcelConfig(string filePath)
        {
            if (filePath == string.Empty)
            {
                throw new Exception("path不能为空！");
            }
            try
            {
                using (StreamReader srFile = File.OpenText(filePath))
                {
                    using (JsonTextReader jsonReader = new JsonTextReader(srFile))
                    {
                        JObject o = (JObject)JToken.ReadFrom(jsonReader);
                        var excelConfig = o.ToObject<ExcelConfig>();
                        this._excelConfig = excelConfig;
                    }
                }
            }
            catch
            {
                throw new Exception("读取配置文件错误，请检查配置文件是否有误");
            }
        }





        #region virtal方法
        /// <summary>
        /// 设置数据行单元格
        /// </summary>
        /// <param name="row">行对象对象</param>
        /// <param name="cell">单元格对象</param>
        /// <param name="field">字段名</param>
        /// <param name="value">值</param>
        public virtual void SetRowData(IRow row,ICell cell,string field,string value)
        {

        }

        /// <summary>
        /// ExcelConfig 文件校验
        /// </summary>
        /// <param name="excelConfig"></param>
        public virtual void ExcelConfigCheck()
        {

        }
        #endregion



        #region 创建1
        /// <summary>
        /// 创建普通表格
        /// </summary>
        /// <param name="formData"></param>
        /// <returns></returns>
        private HSSFWorkbook CreateCommonTable(DataTable formData)
        {
            CreateTableHeader(_excelConfig.TableHeader);
            CreateTableData(formData);
            return this.workbook;
        }

        /// <summary>
        /// 创建普通表格
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <returns></returns>
        private HSSFWorkbook CreateCommonTable<T>(List<T> dataList)
        {
            CreateTableHeader(_excelConfig.TableHeader);
            CreateTableData(dataList);
            return this.workbook;
        }

        /// <summary>
        /// 创建表单表格
        /// </summary>
        /// <param name="excelConfig"></param>
        private HSSFWorkbook CreateFormTable(List<FormOption> formOptions, DataTable formData)
        {
            _formOption = formOptions;
            CreateFormHeader(_excelConfig.FormHeader);
            CreateTableHeader(_excelConfig.TableHeader);
            CreateTableData(formData);
            return this.workbook;
        }
        /// <summary>
        ///  创建表单表格
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="formOptions"></param>
        /// <param name="dataList"></param>
        private HSSFWorkbook CreateFormTable<T>(List<FormOption> formOptions, List<T> dataList)
        {
            _formOption = formOptions;
            CreateFormHeader(_excelConfig.FormHeader);
            CreateTableHeader(_excelConfig.TableHeader);
            CreateTableData(dataList);
            return this.workbook;
        }
        #endregion

        #region 创建表单头
        private void CreateFormHeader(FormHeader formHeader)
        {
            if (formHeader.FormFields == null || formHeader.FormFields.Count == 0)
            {
                throw new Exception("请配置json文件中的‘FormFields’项");
            }
            // //创建单元格 合并单元格
            formHeader.FormFields.ForEach(item =>
            {
                CreateFormCell(item);
            });
        }
        /// <summary>
        /// 创建表单单元格
        /// </summary>
        /// <param name="formItem"></param>
        private void CreateFormCell(FormItem formItem)
        {
            var row = sheet.CreateRow(formItem.RowIndex);
            var textcell = CreateCell(row, formItem);
            ICellStyle textStyle = null;
            if (formItem.CellStyle != null)
            {
                textStyle = CreateCellStyle(formItem.CellStyle);
            }
            else
            {
                textStyle = CreateDefaultCellStyle();
            }
            textcell.CellStyle = textStyle;
            textcell.SetCellType(CellType.String);
            textcell.SetCellValue(formItem.Name);
            int startRow = 0, endRow = 0, startColumn = 0, endColunmn = 0;
            startRow = endRow = formItem.RowIndex;
            startColumn = endColunmn = formItem.ColumnsIndex;
            if (formItem.ColSpace > 0)
            {
                endColunmn = endColunmn + formItem.ColSpace;
            }
            if (formItem.RowSpace > 0)
            {
                endRow = endRow + formItem.RowSpace;
            }
            if (formItem.RowSpace > 0 || formItem.ColSpace > 0)
            {
                MergedRegion(startRow,endRow,startColumn,endColunmn, formItem.HasBorder);
            }
            CreateFormValueCell(row, formItem, startRow, endRow, startColumn, endColunmn);
        }

        /// <summary>
        /// 创建表单value单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="formItem"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColunmn"></param>
        private void CreateFormValueCell(IRow row, FormItem formItem,int startRow, int endRow, int startColumn, int endColunmn)
        {
            ICellStyle valuStyle;
            ICell valueCell = null;
            if (formItem.CellStyle != null)
            {
                valuStyle = CreateCellStyle(formItem.CellStyle);
                valuStyle.FillPattern = FillPattern.NoFill; //不需要与text的颜色填充
                valuStyle.GetFont(workbook).Color = 0;
            }
            else
            {
                valuStyle = CreateDefaultCellStyle();
            }
            if (formItem.ValueCellConfig != null && formItem.ValueCellConfig.CellStyle != null)
            {
                valuStyle = CreateCellStyle(formItem.ValueCellConfig.CellStyle);
            }
            if (formItem.Orientation > 1) //上下结构
            {
                var rowIndex = (formItem.RowSpace == 0 ? formItem.RowIndex + 1 : formItem.RowIndex + formItem.RowSpace + 1);
                var valueRow = sheet.CreateRow(rowIndex);
                valueCell = CreateCell(valueRow, formItem.ColumnsIndex);
                if (formItem.ValueCellConfig != null && formItem.ValueCellConfig.RowSpace > 0)
                {
                    MergedRegion(rowIndex, rowIndex + formItem.ValueCellConfig.RowSpace, startColumn, endColunmn, formItem.HasBorder);
                }else if(formItem.ColSpace > 0 || formItem.RowSpace > 0)
                {
                    MergedRegion(rowIndex, rowIndex, startColumn, endColunmn, formItem.HasBorder);
                }
            }
            else //左右
            {
                var colIndex = (formItem.ColSpace == 0 ? formItem.ColumnsIndex + 1 : formItem.ColumnsIndex + formItem.ColSpace + 1);
                valueCell = CreateCell(row, colIndex);
                if (formItem.ValueCellConfig != null && formItem.ValueCellConfig.ColSpace > 0)
                {
                    MergedRegion(startRow, endRow, colIndex, colIndex + formItem.ValueCellConfig.ColSpace, formItem.HasBorder);
                }
                else if (formItem.ColSpace > 0 || formItem.RowSpace > 0)
                {
                    MergedRegion(startRow, endRow, colIndex, colIndex, formItem.HasBorder);
                }
            }
            valueCell.SetCellType(CellType.String);
            valueCell.CellStyle = valuStyle;
            valueCell.SetCellValue(GetFormFiledValue(formItem));
        }

        #endregion

        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="tableHeader"></param>
        private void CreateTableHeader(TableHeader tableHeader)
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
                row.Height = (short)(tableHeader.Height / 2 * 20);
                row2.Height = (short)(tableHeader.Height / 2 * 20);
            }
            else
            {
                row.Height = (short)(tableHeader.Height * 20);
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
                        MergedRegion(tableHeader.StratRow, tableHeader.StratRow + 1, item.ColumnsIndex, item.ColumnsIndex, tableHeader.HasBorder);
                        sheet.SetColumnWidth(item.ColumnsIndex, item.Width * 256);
                    }

                    totalLen++;
                    if (hasChild)
                    {
                        MergedRegion(tableHeader.StratRow, tableHeader.StratRow, item.ColumnsIndex, item.ColumnsIndex + item.ChildHeaders.Count-1, tableHeader.HasBorder);
                        int childIndex = 0;
                        totalLen = totalLen + (item.ChildHeaders.Count - 1);
                        //绘制子表头
                        item.ChildHeaders.OrderBy(o => o.OrderNum).ToList().ForEach(child =>
                            {
                                child.ColumnsIndex = item.ColumnsIndex + childIndex;
                                var cell1 = CreateCell(row2, child);
                                if (child.CellStyle != null)
                                {
                                    style = CreateCellStyle(child.CellStyle);
                                }
                                cell1.CellStyle = style;
                                cell1.SetCellType(CellType.String);
                                cell1.SetCellValue(child.Name);
                                childIndex++;
                                columConfig.Add(child.Field, child.ColumnsIndex);
                                sheet.SetColumnWidth(child.ColumnsIndex, child.Width * 256);
                            });
                    }
                    else
                    {
                        columConfig.Add(item.Field, item.ColumnsIndex);
                    }
                });
        }

        /// <summary>
        /// datatable 创建数据
        /// </summary>
        /// <param name="formDt"></param>
        private void CreateTableData(DataTable formDt)
        {
            if (formDt == null || formDt.Rows.Count == 0)
            {
                return;
            }
            int startRowIndex = _excelConfig.TableHeader.StratRow + 1;
            int startColumnIndex = _excelConfig.TableHeader.StratColumn;
            if (_excelConfig.TableHeader.Columns.Any(a => a.ChildHeaders != null && a.ChildHeaders.Count > 0))
            {
                startRowIndex++;
            }
            var style = CreateDefaultCellStyle();
            AddAllBorder(style);
            int rowIndex = 0;
            foreach (DataRow dr in formDt.Rows)
            {
                var row = sheet.CreateRow(startRowIndex+ rowIndex);
                row.Height = 20 * 20;
                foreach (var keyPair in columConfig)
                {
                    var cell = CreateCell(row, keyPair.Value);
                    cell.CellStyle = style;
                     var value = dr[keyPair.Key];
                    if (value != null)
                    {
                        cell.SetCellValue(value.ToString());
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                    //供单元格扩展使用
                    SetRowData(row,cell, keyPair.Key, value.ToString());
                }
                rowIndex++;
            }
        }

       

        /// <summary>
        /// list创建数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        private void CreateTableData<T>(List<T> list)
        {
            if (list == null || list.Count == 0)
            {
                return;
            }
            int startRowIndex = _excelConfig.TableHeader.StratRow + 1;
            int startColumnIndex = _excelConfig.TableHeader.StratColumn;
            if (_excelConfig.TableHeader.Columns.Any(a => a.ChildHeaders != null && a.ChildHeaders.Count > 0))
            {
                startRowIndex++;
            }
            var style = CreateDefaultCellStyle();
            AddAllBorder(style);
            int rowIndex = 0;
            list.ForEach(item =>
            {
                var row = sheet.CreateRow(startRowIndex + rowIndex);
                row.Height = 20 * 20;
                foreach (var keyPair in columConfig)
                {
                    var cell = CreateCell(row, keyPair.Value);
                    cell.CellStyle = style;
                    object value = null;
                    Type t = typeof(T);  //创建一个T类型的类型
                    var propertys = t.GetProperties();
                    foreach (var pro in propertys)
                    {
                        if (pro.Name == keyPair.Key)
                        {
                            value = pro.GetValue(item, null);//获取属性的值
                        }
                    }
                    if (value != null)
                    {
                        cell.SetCellValue(value.ToString());
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                    //供单元格扩展使用
                    SetRowData(row, cell, keyPair.Key, value.ToString());
                }
                rowIndex++;
            });
        }


        #region 公共

        /// <summary>
        /// 表单获取字段值
        /// </summary>
        /// <param name="baseCell"></param>
        /// <returns></returns>
        private string GetFormFiledValue(BaseCell baseCell)
        {
            if (!string.IsNullOrEmpty(baseCell.FixedValue.ToString()))
            {
                return baseCell.FixedValue.ToString();
            }
            if (this._formOption == null)
            {
                return "";
            }
            var option=  this._formOption.Where(w => w.FildName == baseCell.Field).FirstOrDefault();
            if (option != null)
            {
                return option.FildValue;
            }
            return "";
        }


        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="mergedConfig"></param>
        /// <param name="hasBorder"></param>
        private void MergedRegion(int startRowIndex, int endRowIndex, int startCoulumnIndex, int endCoulumnIndex, bool hasBorder)
        {
            var region = new CellRangeAddress(startRowIndex, endRowIndex, startCoulumnIndex, endCoulumnIndex);
            sheet.AddMergedRegion(region);
            if (hasBorder)
            {
                AddBorderRegion(region);
            }
        }

        /// <summary>
        /// 合并区域添加边框
        /// </summary>
        /// <param name="region"></param>
        private void AddBorderRegion(CellRangeAddress region)
        {
            RegionUtil.SetBorderBottom(1, region, sheet, workbook);
            RegionUtil.SetBorderLeft(1, region, sheet, workbook);
            RegionUtil.SetBorderRight(1, region, sheet, workbook);
            RegionUtil.SetBorderTop(1, region, sheet, workbook);
        }


        /// <summary>
        /// 检查颜色是否
        /// </summary>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        private bool CheckColor(short fontColor)
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
        private ICellStyle CreateCellStyle(CellStyle cellStyle)
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
        private ICellStyle CreateHeaderStyle(TableHeader tableHeader)
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
        private ICell CreateCell(IRow row, BaseCell baseCell)
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
        private ICell CreateCell(IRow row, int cIndex)
        {

            return row.CreateCell(cIndex);
        }

        #endregion









    }
}
