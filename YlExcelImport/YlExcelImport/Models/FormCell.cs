using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    public class FormCell : BaseCell
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
        /// 结构（1：左右(默认)，2：上下）
        /// </summary>
        public int Orientation { get; set; } = 1;

        public BaseCell PreCell { get; set; }
    }
}
