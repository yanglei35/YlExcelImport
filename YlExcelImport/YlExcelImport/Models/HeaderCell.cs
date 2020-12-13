using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
   public  class HeaderCell :BaseCell
    {
        /// <summary>
        /// 排序号
        /// </summary>
        public int OrderNum { get; set; }

        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; } = 18;

        /// <summary>
        /// 嵌套列表
        /// </summary>
        public List<HeaderCell> ChildHeaders { get; set; }
    }
}
