using EasyNPOI.Attributes;
using EasyNPOI.Enums;
using EasyNPOI.Models.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Services
{
    public class NPOIExcelHelper
    {
        public byte[] Export<T>(ExportOption<T> exportOption) where T : class, new()
        {
            IWorkbook workbook = null;
            if (exportOption.ExportType == ExportType.XLS)
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                workbook = new XSSFWorkbook();
            }

            ISheet sheet = workbook.CreateSheet(exportOption.SheetName);

            var headerDict = ExportMappingDictFactory.CreateInstance(typeof(T));

            SetHeader<T>(sheet, exportOption.HeaderRowIndex, headerDict);

            if (exportOption.Data != null && exportOption.Data.Count > 0)
            {
                SetDataRows(sheet, exportOption.DataRowStartIndex, exportOption.Data, headerDict);
            }

            return workbook?.ToBytes();
        }

        private void SetHeader<T>(ISheet sheet, int headerRowIndex, Dictionary<string, string> headerDict) where T : class, new()
        {
            IRow row = sheet.CreateRow(headerRowIndex);
            int colIndex = 0;
            foreach (var kvp in headerDict)
            {
                row.CreateCell(colIndex).SetCellValue(kvp.Value);
                colIndex++;
            }
        }

        private void SetDataRows<T>(ISheet sheet, int dataRowStartIndex, List<T> datas, Dictionary<string, string> headerDict) where T : class, new()
        {
            if (datas.Count <= 0) return;

            for (var i = 0; i < datas.Count; i++)
            {
                int colIndex = 0;
                IRow row = sheet.CreateRow(dataRowStartIndex + i);
                T dto = datas[i];

                foreach (var kvp in headerDict)
                {
                    row.CreateCell(colIndex).SetCellValue(dto.GetStringValue(kvp.Key));
                    colIndex++;
                }
            }
        }

        /// <summary>
        /// 简单实现图片替换
        /// </summary>
        /// <param name="fileFullPath"></param>
        /// <param name="replacements"></param>
        public void ReplacePlaceholders(string fileFullPath, List<EasyNPOI.Models.Excel.Replacement> replacements)
        {
            if (string.IsNullOrEmpty(fileFullPath))
            {
                throw new ArgumentNullException();
            }

            var fs = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read);
            var extName = Path.GetExtension(fileFullPath);

            IWorkbook wb;
            if (extName == ".xls")
            {
                wb = new HSSFWorkbook(fs);
            }
            else if (extName == ".xlsx")
            {
                wb = new XSSFWorkbook(fs);
            }
            else
            {
                wb = null;
            }
            if (wb == null)
            {
                throw new ArgumentNullException();
            }

            var snum = wb.NumberOfSheets;
            var reSave = false;
            for (int i = 0; i < snum; i++)
            {
                //if (wb.IsSheetHidden(i))
                //{
                //    continue;
                //}

                var sheet = wb.GetSheetAt(i);
                IRow row = sheet.GetRow(sheet.LastRowNum);
                if (row == null)
                {
                    continue;
                }
                var cells = row.Cells.Where(p => p.CellType == CellType.String).ToList();
                foreach (var x in cells)
                {
                    var colNum = x.ColumnIndex;
                    var rowNum = x.RowIndex;
                    var cellText = x.StringCellValue;
                    if (string.IsNullOrEmpty(cellText))
                    {
                        continue;
                    }

                    foreach (var replace in replacements)
                    {
                        if (cellText.Contains(replace.Placeholder))
                        {
                            //获取指定图片的字节流
                            //byte[] bytes = new byte[pictureData.Length];
                            //pictureData.Read(bytes, 0, bytes.Length);
                            byte[] bytes = System.IO.File.ReadAllBytes(replace.PictureUrl);

                            //将图片添加到工作簿中，返回值为该图片在工作表中的索引（从0开始）
                            //图片所在工作簿索引理解：如果原Excel中没有图片，那执行下面的语句后，该图片为Excel中的第1张图片，其索引为0；
                            //同理，如果原Excel中已经有1张图片，执行下面的语句后，该图片为Excel中的第2张图片，其索引为1；
                            int pictureIdx = wb.AddPicture(bytes, NPOI.SS.UserModel.PictureType.PNG);

                            if (extName == ".xls")
                            {
                                HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, colNum, rowNum, colNum + 1, rowNum + 1);
                                HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                                reSave = true;
                            }
                            else if (extName == ".xlsx")
                            {
                                //创建画布
                                XSSFDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                                //设置图片坐标与大小
                                //函数原型：XSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)；
                                //坐标(col1,row1)表示图片左上角所在单元格的位置，均从0开始，比如(5,2)表示(第五列，第三行),即F3；注意：图片左上角坐标与(col1,row1)单元格左上角坐标重合
                                //坐标(col2,row2)表示图片右下角所在单元格的位置，均从0开始，比如(10,3)表示(第十一列，第四行),即K4；注意：图片右下角坐标与(col2,row2)单元格左上角坐标重合
                                //坐标(dx1,dy1)表示图片左上角在单元格(col1,row1)基础上的偏移量(往右下方偏移)；(dx1，dy1)的最大值为(1023, 255),为一个单元格的大小
                                //坐标(dx2,dy2)表示图片右下角在单元格(col2,row2)基础上的偏移量(往右下方偏移)；(dx2,dy2)的最大值为(1023, 255),为一个单元格的大小
                                //注意：目前测试发现，对于.xlsx后缀的Excel文件，偏移量设置(dx1,dy1)(dx2,dy2)无效；只会对.xls生效
                                XSSFClientAnchor anchor = new XSSFClientAnchor(100, 100, 0, 0, colNum, rowNum, colNum + 1, rowNum + 1);
                                //正式在指定位置插入图片
                                XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                                reSave = true;
                            }

                            cellText = cellText.Replace(replace.Placeholder, "");
                            x.SetCellValue(cellText);
                        }

                    }
                }

            }

            if (reSave)
            {
                //创建一个新的Excel文件流，可以和原文件名不一样，
                //如果不一样，则会创建一个新的Excel文件；如果一样，则会覆盖原文件
                FileStream file = new FileStream(fileFullPath, FileMode.Create);
                //将已插入图片的Excel流写入新创建的Excel中
                wb.Write(file);
                //关闭工作簿
                wb.Close();
            }
        }


    }


    public static class ExcelExtensions
    {

        /// <summary>
        /// 将IWorkbook转换为byte数组
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this IWorkbook workbook)
        {
            byte[] result;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                result = ms.ToArray();
            }

            return result;
        }

        /// <summary>
        /// 反射获取导出DTO某个属性的值
        /// </summary>
        /// <param name="export"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string GetStringValue<T>(this T export, string propertyName)
        {
            string strVal = string.Empty;
            var prop = export.GetType().GetProperties().Where(p => p.Name.Equals(propertyName)).SingleOrDefault();
            if (prop != null)
            {
                strVal = prop.GetValue(export) == null ? string.Empty : prop.GetValue(export).ToString();
            }

            return strVal;
        }


        #region 获取excel中的图片信息
        public class PicturesInfo
        {
            public int MinRow { get; set; }
            public int MaxRow { get; set; }
            public int MinCol { get; set; }
            public int MaxCol { get; set; }
            public Byte[] PictureData { get; private set; }

            public PicturesInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData)
            {
                this.MinRow = minRow;
                this.MaxRow = maxRow;
                this.MinCol = minCol;
                this.MaxCol = maxCol;
                this.PictureData = pictureData;
            }
        }

        public static List<PicturesInfo> GetAllPictureInfos(this ISheet sheet)
        {
            return sheet.GetAllPictureInfos(null, null, null, null);
        }

        public static List<PicturesInfo> GetAllPictureInfos(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                return GetAllPictureInfos((HSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else if (sheet is XSSFSheet)
            {
                return GetAllPictureInfos((XSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else
            {
                throw new Exception("未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！");
            }
        }

        private static List<PicturesInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                var shapeList = shapeContainer.Children;
                foreach (var shape in shapeList)
                {
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture)shape;
                        var anchor = (HSSFClientAnchor)shape.Anchor;

                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                        {
                            picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static List<PicturesInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var documentPartList = sheet.GetRelations();
            foreach (var documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing)documentPart;
                    var shapeList = drawing.GetShapes();
                    foreach (var shape in shapeList)
                    {
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture)shape;
                            var anchor = picture.GetPreferredSize();

                            if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                            {
                                picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data));
                            }
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
        {
            int _rangeMinRow = rangeMinRow ?? pictureMinRow;
            int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
            int _rangeMinCol = rangeMinCol ?? pictureMinCol;
            int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

            if (onlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            else
            {
                return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
            }
        }
        #endregion

    }


    /// <summary>
    /// 映射字典工厂
    /// </summary>
    public static class ExportMappingDictFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 创建映射字段，键：DTO属性名，值：Excel列名
        /// </summary>
        /// <param name="exportType"></param>
        /// <returns></returns>
        public static Dictionary<string, string> CreateInstance(Type exportType)
        {
            var key = exportType;
            if (Table[key] != null)
            {
                return (Dictionary<string, string>)Table[key];
            }

            Dictionary<string, string> dict = new Dictionary<string, string>();
            exportType.GetProperties().ToList().ForEach(p =>
            {
                if (p.IsDefined(typeof(ColNameAttribute)))
                {
                    dict.Add(p.Name, p.GetCustomAttribute<ColNameAttribute>().ColName);
                }
                else
                {
                    dict.Add(p.Name, p.Name);
                }
            });

            Table[key] = dict;
            return dict;
        }

    }
}
