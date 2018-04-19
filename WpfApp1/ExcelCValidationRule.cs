using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelSampleApp
{
    class ExcelCValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            //ExcelManager _manager = new ExcelManager();
            //List<ExcelB> listB = _manager.LoadCollectionExcelB();

            ExcelC entry = (value as BindingGroup).Items[0] as ExcelC;
            if (string.IsNullOrEmpty(entry.Val1))
            {
                return new ValidationResult(false, "Value 1 must not be empty!");
            }


            //bool contains = false;

            //foreach (var item in listB)
            //{
            //    if (item.Val3Main.Equals(entry.Val3))
            //        contains = true;
            //}

            //if (contains)
            //    return ValidationResult.ValidResult;
            //else
            //    return new ValidationResult(false, "Value 3 must exist in ExcelB!");

            return ValidationResult.ValidResult;
        }
    }
}
