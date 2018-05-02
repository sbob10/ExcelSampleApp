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
    /// <summary>
    /// Interaktionslogik für EditValueExcelCWindow.xaml
    /// </summary>
    public partial class EditValueExcelCWindow : Window
    {

        EditValueExcelCWindowViewModel vm;
        Action<ExcelC> callbackFromParentWindow;

        public EditValueExcelCWindow(Action<ExcelC> editingAction, ExcelC entryToEdit)
        {
            InitializeComponent();

            vm = (EditValueExcelCWindowViewModel)this.DataContext;
            vm.EditExcelCValue = entryToEdit;
            vm.CloseDialogFunc = new Func<int, string>(ShowEditEntryDialogFuncExcelC);

            callbackFromParentWindow = editingAction;
        }

        private void PassEntryExcelCToParentWindowCallback(object sender, EventArgs e)
        {
            callbackFromParentWindow(vm.EditExcelCValue);
        }

        private String ShowEditEntryDialogFuncExcelC(int statusCode)
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
