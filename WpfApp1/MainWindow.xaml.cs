using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using SpreadsheetLight;

namespace ExcelSampleApp
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        MainWindowViewModel vm;


        public MainWindow()
        {
            InitializeComponent();
            vm = (MainWindowViewModel) this.DataContext;
            vm.ShowAddEntryDialogFuncExcelC = new Func<int, int, string>(ShowAddEntryDialogFuncExcelC);
        }

        protected override void OnClosing(CancelEventArgs e)
        {

            vm.StoreAllCollections();

            var result = MessageBox.Show("Do you want to cancel?", "Leaving", MessageBoxButton.YesNo, MessageBoxImage.Hand);
            if (result == MessageBoxResult.Yes)
            {
                base.OnClosing(e);
                return;
            }
            e.Cancel = true;
        }

        private string ShowAddEntryDialogFuncExcelC(int pseudoValue1, int pseudoValue2)
        {
            var dialog = new AddValueExcelCWindow(AddEntryFromDialogExcelC);
            dialog.ShowDialog();

            return ""; //just pseudo
        }

        private void AddEntryFromDialogExcelC(ExcelC entry)
        {
            vm.AddEntryExcelCFromCodeBehind(entry);
        }

    }
}
