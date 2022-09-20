using System;
using System.Collections.Generic;
using System.Text;

namespace NightlyRouteToSlack 
{
    class Spreadsheet
    {
        public SpreadSheetRow HeaderRow { get; set; }
        public List<SpreadSheetRow> Rows { get; set; }
    }
}
