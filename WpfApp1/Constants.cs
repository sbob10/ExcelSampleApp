using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSampleApp
{
    static class Constants
    {
        public static readonly List<String> Months = new List<String>(new String[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" });

        public static String getCurrentDateMonth()
        {
            return DateTime.Now.ToString("MMMM");
        } 
    }
}
