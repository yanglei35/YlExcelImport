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
        /// 排列顺序
        /// </summary>
        public int OrderNum { get; set; }
        /// <summary>
        /// 嵌套列表
        /// </summary>
        public List<BaseCell> ChildHeaders { get; set; }
    }
}
