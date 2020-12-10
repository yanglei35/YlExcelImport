using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    public class ExcelConfig
    {
        /// <summary>
        /// excel名称
        /// </summary>
        public string ExcelName { get; set; }

        /// <summary>
        /// excel导出类型
        /// </summary>
        public int ExcelType { get; set; }

        /// <summary>
        /// 表头
        /// </summary>
        public FormHeader FormHeader { get; set; }

        /// <summary>
        /// 表格表头
        /// </summary>
        public TableHeader TableHeader { get; set; }

    }
}
