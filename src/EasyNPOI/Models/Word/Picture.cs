﻿using EasyNPOI.Enums;
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
        /// 图片绝对地址（如果PictureData不为空则不用传）
        /// </summary>
        public string PictureUrl { get; set; }

        /// <summary>
        /// 图片类型，默认PNG
        /// </summary>
        public PictureTypeEnum PictureType { get; set; } = PictureTypeEnum.PNG;

        /// <summary>
        /// 文件名，默认“picture”
        /// </summary>
        public string FileName { get; set; } = "picture";

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