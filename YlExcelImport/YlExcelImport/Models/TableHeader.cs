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
        /// 背景色
        /// </summary>
        public short BackGroundColor { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public short FontColor { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public short FontSize { get; set; }

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; } = 15;

        /// <summary>
        /// 字段集合
        /// </summary>
        public List<HeaderCell> Columns { get; set; }
    }
}
