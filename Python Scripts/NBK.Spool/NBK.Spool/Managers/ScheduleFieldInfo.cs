using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Element = Autodesk.Revit.DB.Element;

namespace NBK.Spool
{
    public class ScheduleFieldInfo
    {
        public string ColumnHeader { get; set; }
        public string ParameterName { get; set; }
        public string Unit { get; set; }
        public Parameter Parameter { get; set; }
        public double ColumWidthFeet { get; set; }

        public ScheduleFieldInfo(){}

        public ScheduleFieldInfo(string columnHeader, string parameterName, double columnWidthInches)
        {
            this.ColumnHeader = columnHeader;
            this.ParameterName = parameterName;
            this.ColumWidthFeet = columnWidthInches / 12.0;
        }
    }
}
