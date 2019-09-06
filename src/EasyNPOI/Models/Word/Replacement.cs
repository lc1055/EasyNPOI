using EasyNPOI.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Models.Word
{
    /// <summary>
    /// 占位符替换对象(文本/图片)
    /// </summary>
    public class ReplacementBasic
    {
        /// <summary>
        /// 占位符
        /// </summary>
        public string Placeholder { get; set; }

        /// <summary>
        /// 替换的文本
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// 替换的图片
        /// </summary>
        public IEnumerable<Picture> Pictures { get; set; }

        /// <summary>
        /// 占位符类型
        /// </summary>
        public PlaceholderTypeEnum Type { get; set; }
    }

    /// <summary>
    /// 占位符替换对象(表格)
    /// </summary>
    public class ReplacementGrid
    {
        /// <summary>
        /// 占位符
        /// </summary>
        public string Placeholder { get; set; }

        /// <summary>
        /// 替换的列表数据
        /// </summary>
        public List<ReplacementRow> Rows { get; set; } = new List<ReplacementRow>();

        /// <summary>
        /// 占位符类型
        /// </summary>
        public PlaceholderTypeEnum Type { get { return PlaceholderTypeEnum.Grid; } }
    }

    public class ReplacementRow
    {
        public List<ReplacementBasic> Cells { get; set; } = new List<ReplacementBasic>();
    }

    

}
