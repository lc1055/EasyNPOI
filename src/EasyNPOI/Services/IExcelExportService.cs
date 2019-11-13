using EasyNPOI.Models.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Services
{
    public interface IExcelExportService
    {
        Task<byte[]> ExportAsync<T>(ExportOption<T> exportOption) where T : class, new();
    }
}
