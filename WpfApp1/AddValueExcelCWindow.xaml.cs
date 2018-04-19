using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelSampleApp
{
   
    public partial class AddValueExcelCWindow : Window
    {
        AddValueExcelCWindowViewModel vm;
        Action<ExcelC> callbackFromParentWindow;
        // Tutorial: https://stackoverflow.com/questions/8590055/pass-from-child-window-to-parent-window

        public AddValueExcelCWindow(Action<ExcelC> addingAction)
        {
            InitializeComponent();
            vm = (AddValueExcelCWindowViewModel)this.DataContext;
            vm.CloseDialogFunc = new Func<int, string>(ShowAddEntryDialogFuncExcelC);

            callbackFromParentWindow = addingAction;
        }

        private void PassEntryExcelCToParentWindowCallback(object sender, EventArgs e)
        {
            callbackFromParentWindow(vm.AddExcelCValue);
        }

        private String ShowAddEntryDialogFuncExcelC(int statusCode)
        {
            if (statusCode == 1)
            {
                Closed += PassEntryExcelCToParentWindowCallback;
            }
            this.Close();

            return ""; //just pseudo
        }

    }
}
