using EasyNPOI.Attributes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTest
{
    [TestClass]
    public class ExcelTest
    {
        [TestMethod]
        public async Task ExcelExportTest()
        {
            string dir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(dir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

            var carDTO = new ExcelCarTemplateDTO()
            {
                Age = 10,
                CarCode = "鄂A123456",
                IdentityNumber = "test",
                Mobile = "test",
                Name = "test",
                RegisterDate = DateTime.Now
            };

            var list = new List<ExcelCarTemplateDTO>();

            for (int i = 0; i < 10; i++)
            {
                list.Add(carDTO);
            }

            EasyNPOI.Services.ExcelExportService srv = new EasyNPOI.Services.ExcelExportService();
            var bytes = await srv.ExportAsync(new EasyNPOI.Models.Excel.ExportOption<ExcelCarTemplateDTO> { Data = list });
            
            File.WriteAllBytes(fileUrl, bytes);
        }


        /// <summary>
        /// 替换excel中的占位符为图片，目前替换后的图片只能占满整格，使用场景有限。
        /// 当然，如果是纯文字替换的话，是没有问题的。修改Replacement类中的属性和相关代码就行了。
        /// </summary>
        [TestMethod]
        public void TestReplaceExcelStr()
        {
            var path = @"d:\123\111.xlsx";
            var placeholderList = new List<EasyNPOI.Models.Excel.Replacement>
            {
                new EasyNPOI.Models.Excel.Replacement{ Placeholder="{ss}", PictureUrl=@"d:\123\heihei.png" },
                new EasyNPOI.Models.Excel.Replacement{ Placeholder="{tt}", PictureUrl=@"d:\123\heihei.png" },
            };

            EasyNPOI.Services.ExcelExportService srv = new EasyNPOI.Services.ExcelExportService();
            srv.ReplaceAsync(path, placeholderList);

        }

    }


    public class ExcelCarTemplateDTO
    {
        [ColName("车牌号")]
        public string CarCode { get; set; }

        [ColName("手机号")]
        public string Mobile { get; set; }

        [ColName("身份证号")]
        public string IdentityNumber { get; set; }

        [ColName("姓名")]
        public string Name { get; set; }

        [ColName("注册日期")]
        public DateTime RegisterDate { get; set; }

        [ColName("年龄")]
        public int Age { get; set; }
    }



}
