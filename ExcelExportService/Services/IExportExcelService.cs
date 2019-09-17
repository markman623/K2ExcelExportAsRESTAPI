using ExcelExportService.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelExportService.Services
{
    public interface IExportExcelService
    {
        Task<ExcelExport> ExportToExcelAsync(string smoQuery, string fileName);
    }
}
