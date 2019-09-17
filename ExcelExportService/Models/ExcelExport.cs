using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace ExcelExportService.Models
{
    public class ExcelExport
    {
        public string FileName { get; set; }

        public string ExcelFile { get; set; }

        public int Size { get; set; }
    }
}
