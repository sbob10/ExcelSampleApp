using DevExpress.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelSampleApp
{
    class AddValueExcelCWindowViewModel : ViewModelBase
    {

        #region Properties
        private ExcelC _addExcelCValue;

        public ExcelC AddExcelCValue
        {
            get { return _addExcelCValue; }
            set { SetProperty(ref _addExcelCValue, value, () => AddExcelCValue); }
        }
        #endregion Properties

        #region Commands&Services

        public Func<int, string> CloseDialogFunc { get; set; }

        public ICommand AddEntryExcelCCommand { get; set; }
        public ICommand CloseDialogCommand { get; set; }

        public IMessageBoxService MessageBoxService { get { return ServiceContainer.GetService<IMessageBoxService>(); } set { } }
        public IServiceContainer IServiceContainer { get; set; }

        #endregion Commands

        #region Initialisations

        public AddValueExcelCWindowViewModel()
        {
            InitCommandsAndServices();
            InitValuesAndCollections();
        }

        private void InitValuesAndCollections()
        {
            AddExcelCValue = new ExcelC();
        }

        private void InitCommandsAndServices()
        {
            AddEntryExcelCCommand = new DelegateCommand(AddEntryExcelC);
            CloseDialogCommand = new DelegateCommand(CloseDialog);

            IServiceContainer = new ServiceContainer(this);
            ServiceContainer.RegisterService(new DevExpress.Mvvm.UI.MessageBoxService());
        }

        #endregion Initialisations

        #region Private Methods
    


        #endregion Private Methods

        #region Command Methods

        private void AddEntryExcelC()
        {
            if (string.IsNullOrEmpty(AddExcelCValue.Val1))
                MessageBoxService.ShowMessage("Value 1 must not be empty!", "Error", MessageButton.OK, MessageIcon.Information);
            else
                CloseDialogFunc.Invoke(1);
        }

        private void CloseDialog()
        {
            CloseDialogFunc.Invoke(0);
        }

        #endregion Command Methods

    }
}
