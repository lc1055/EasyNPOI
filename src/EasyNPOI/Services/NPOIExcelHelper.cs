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
