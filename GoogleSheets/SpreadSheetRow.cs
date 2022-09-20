using System;
using System.Collections.Generic;

namespace NightlyRouteToSlack 
{
    public class SpreadSheetRow
    {
        private readonly IList<Object> _values;
        public SpreadSheetRow(IList<Object> values)
        {
            _values = values;
        }

        public string ColumnValue(int columnIndex) => _getValue(columnIndex);

        private string _getValue(int columnIndex)
        {
            try
            {
                var s = _values[columnIndex].ToString();
                return s;
            }
            catch (Exception ex)
            {
                return String.Empty;
            }
        }
    }
}