using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    public class TableHeader
    {
        /// <summary>
        /// 开始行
        /// </summary>
        public int StratRow { get; set; }

        /// <summary>
        /// 开始列
        /// </summary>
        public int StratColumn { get; set; }

        /// <summary>
        /// 是否有边框
        /// </summary>
        public bool HasBorder { get; set; }

        /// <summary>
        /// 字段集合
        /// </summary>
        public List<HeaderCell> Columns { get; set; }
    }
}
