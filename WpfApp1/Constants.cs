using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSampleApp
{
    static class Constants
    {
        public static readonly List<String> Months = new List<String>(new String[] { "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember" });

        public static String getCurrentDateMonth()
        {
            return DateTime.Now.ToString("MMMM");
        } 
    }
}
