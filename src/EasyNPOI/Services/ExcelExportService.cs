using EasyNPOI.Enums;
using EasyNPOI.Models.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Services
{
    public class ExcelExportService : IExcelExportService
    {
        public Task<byte[]> ExportAsync<T>(ExportOption<T> exportOption) where T : class, new()
        {
            if (exportOption.ExportType == Enums.ExportType.XLS && exportOption.Data.Count > 65535)
            {
                throw new InvalidOperationException("xls格式文件最多支持65536行数据");
            }

            if (exportOption.ExportType == Enums.ExportType.XLSX && exportOption.Data.Count > 1048575)
            {
                throw new InvalidOperationException("xlsx格式文件最多支持1048575行数据");
            }

            NPOIExcelHelper helper = new NPOIExcelHelper();
            var workbookBytes = helper.Export(exportOption);

            //返回byte数组
            return Task.FromResult(workbookBytes);
        }

    }


}
