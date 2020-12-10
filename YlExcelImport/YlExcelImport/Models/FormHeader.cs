using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    /// <summary>
    /// 表单表格表头
    /// </summary>
    public class FormHeader
    {
        /// <summary>
        /// 开始行
        /// </summary>
        public int StratRow { get; set; }

        /// <summary>
        /// 字段集合
        /// </summary>
        public List<FormCell> FormFields { get; set; }
        
    }
}
