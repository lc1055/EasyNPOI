using System;
using System.Collections;
using System.Collections.Generic;
using EasyNPOI.Attributes;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest
{
    [TestClass]
    public class WordTest
    {
        [TestMethod]
        public void WordExportTest()
        {
            string path = Environment.CurrentDirectory;
            string tmplPath = System.IO.Path.Combine(path, "resource", "template.docx");
            string picturePath = System.IO.Path.Combine(path, "resource", "1.png");
            string savePath = System.IO.Path.Combine(path, "resource", "export.docx");


            EasyNPOI.Services.WordExportService srv = new EasyNPOI.Services.WordExportService();


            var car = new Purchase
            {
                User = "马化腾",
                Date = DateTime.Now,
                Note = "老司机开车",
                Check = new EasyNPOI.Models.Word.Picture { PictureUrl = picturePath },
                Approve = new List<EasyNPOI.Models.Word.Picture>
                {
                    new EasyNPOI.Models.Word.Picture { PictureUrl = picturePath, },
                    new EasyNPOI.Models.Word.Picture { PictureUrl = picturePath, }
                },
                Items = new List<Item>
                {
                    new Item { Name = "宝马", Spec = "335", Quantity = 1, Picture = new EasyNPOI.Models.Word.Picture { PictureUrl = picturePath } },
                    new Item { Name = "奔驰", Spec = "glk", Quantity = 2, Picture = new EasyNPOI.Models.Word.Picture { PictureUrl = picturePath } },
                    new Item { Name = "奥迪", Spec = "a8", Quantity = 3, Picture = new EasyNPOI.Models.Word.Picture { PictureUrl = picturePath } },
                },
                Suppliers = new List<Supplier>
                {
                    new Supplier { Name = "京东", ContactUser = "刘强东", PhoneNumber = "12345678" },
                    new Supplier { Name = "淘宝", ContactUser = "马云", PhoneNumber = "87654321" },
                }
            };

            

            srv.ExportByTemplate(tmplPath, savePath, car);
        }

        public class Purchase
        {
            public string User { get; set; }

            public DateTime Date { get; set; }

            public string Note { get; set; }

            [Placeholder("Checker")]
            public EasyNPOI.Models.Word.Picture Check { get; set; }

            [Placeholder("Approver")]
            public List<EasyNPOI.Models.Word.Picture> Approve { get; set; }

            [PlaceholderGrid]
            public List<Item> Items { get; set; }

            [PlaceholderGrid]
            public List<Supplier> Suppliers { get; set; }
        }


        public class Item
        {
            public string Name { get; set; }
            public string Spec { get; set; }          
            public int Quantity { get; set; }
            public EasyNPOI.Models.Word.Picture Picture { get; set; }

        }

        public class Supplier
        {
            public string Name { get; set; }
            public string ContactUser { get; set; }
            public string PhoneNumber { get; set; }
        }
    }
}
