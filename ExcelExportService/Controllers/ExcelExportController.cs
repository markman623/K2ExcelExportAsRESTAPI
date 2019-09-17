using ExcelExportService.Models;
using ExcelExportService.Services;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.Extensions.Options;

namespace ExcelExportService.Controllers
{
    [Route("/[controller]")]
    [ApiController]
    public class ExcelExportController : ControllerBase
    {

        private readonly IExportExcelService _exportExcelService;

        public ExcelExportController(
            IExportExcelService exportExcelService
           
            )
        {
            _exportExcelService = exportExcelService;
        }

        // POST /ExcelExport
        [HttpPost(Name = nameof(ExportFromQuery))]
        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        public async Task<ActionResult<ExcelExport>> ExportFromQuery (
            [FromQuery] string smoName,
            [FromQuery] string fileName)
        {

            var excel = await _exportExcelService.ExportToExcelAsync(smoName, fileName);

            return excel;
        }

    }
}
