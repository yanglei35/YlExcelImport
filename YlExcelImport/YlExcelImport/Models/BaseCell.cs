using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    public class BaseCell
    {
        /// <summary>
        /// 占据几列
        /// </summary>
        public int ColSpace { get; set; }

        /// <summary>
        /// 占据几行
        /// </summary>
        public int RowSpace { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 字段名
        /// </summary>
        public string Filed { get; set; }

        /// <summary>
        /// 固定值
        /// </summary>
        public object FixedValue { get; set; }

        /// <summary>
        /// 开始行
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 开始列
        /// </summary>
        public int ColumnsIndex { get; set; }

        /// <summary>
        /// 单元格样式
        /// </summary>
        public CellStyle CellStyle { get; set; }
    }
}
