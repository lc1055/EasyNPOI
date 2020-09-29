using EasyNPOI.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Models.Word
{
    public class Picture
    {
        /// <summary>
        /// 图片流数据
        /// </summary>
        public Stream PictureData { get; set; }

        /// <summary>
        /// 图片绝对地址（后续程序会将此路径文件转为 <see cref="PictureData"/>）
        /// </summary>
        public string PictureUrl { get; set; }

        /// <summary>
        /// 图片类型，默认PNG
        /// </summary>
        public PictureTypeEnum PictureType { get; set; } = PictureTypeEnum.PNG;

        /// <summary>
        /// 图片文件名。若设置了，则当图片不存在时，会显示此文本
        /// </summary>
        public string FileName { get; set; } = "";

        /// <summary>
        /// 图片宽度，单位厘米，默认14
        /// </summary>
        public decimal Width { get; set; } = 1;

        /// <summary>
        /// 图片高度，单位厘米，默认8
        /// </summary>
        public decimal Height { get; set; } = 1;
    }
}
