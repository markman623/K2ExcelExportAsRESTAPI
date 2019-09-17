using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelExportService
{
    public class K2Context
    {
        public string K2OdataEndpointUrl { get; set; }

        public string K2User { get; set; }

        public string Password { get; set; }

        public string AuthType { get; set; }
    }
}
