using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YlExcelImport.Models
{
    public class CellStyle
    {
        /// <summary>
        /// 字体大小
        /// </summary>
        public int FontSize { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public short FontColor { get; set; }

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        public short BackGroundColor { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 边框 0:无边框(默认)， 1:全边框，2:自定义
        /// </summary>
        public int Border { get; set; }

        /// <summary>
        /// 上边框
        /// </summary>
        public bool BorderTop { get; set; }

        /// <summary>
        /// 下边框
        /// </summary>
        public bool BorderBottom { get; set; }

        /// <summary>
        /// 左边框
        /// </summary>
        public bool BorderLeft { get; set; }

        /// <summary>
        /// 右边框
        /// </summary>
        public bool BorderRight { get; set; }

    }
}
