using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    public class FormItem : BaseCell
    {


        /// <summary>
        /// 结构（1：左右(默认)，2：上下）
        /// </summary>
        public int Orientation { get; set; } = 1;

        /// <summary>
        /// 值
        /// </summary>
        public BaseCell ValueCellConfig { get; set; }
    }
}
