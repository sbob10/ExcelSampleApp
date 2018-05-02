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
    class EditValueExcelCWindowViewModel : ViewModelBase
    {

        #region Properties
        private ExcelC _editExcelCValue;

        public ExcelC EditExcelCValue
        {
            get { return _editExcelCValue; }
            set { SetProperty(ref _editExcelCValue, value, () => EditExcelCValue); }
        }
        #endregion Properties

        #region Commands&Services

        public Func<int, string> CloseDialogFunc { get; set; }

        public ICommand EditEntryExcelCCommand { get; set; }
        public ICommand CloseDialogCommand { get; set; }

        public IMessageBoxService MessageBoxService { get { return ServiceContainer.GetService<IMessageBoxService>(); } set { } }
        public IServiceContainer IServiceContainer { get; set; }

        #endregion Commands

        #region Initialisations

        public EditValueExcelCWindowViewModel()
        {
            InitCommandsAndServices();
            InitValuesAndCollections();
        }

        private void InitValuesAndCollections()
        {
            EditExcelCValue = new ExcelC();
        }

        private void InitCommandsAndServices()
        {
            EditEntryExcelCCommand = new DelegateCommand(EditEntryExcelC);
            CloseDialogCommand = new DelegateCommand(CloseDialog);

            IServiceContainer = new ServiceContainer(this);
            ServiceContainer.RegisterService(new DevExpress.Mvvm.UI.MessageBoxService());
        }

        #endregion Initialisations

        #region Private Methods



        #endregion Private Methods

        #region Command Methods

        private void EditEntryExcelC()
        {
            CloseDialogFunc.Invoke(1);
        }

        private void CloseDialog()
        {
            CloseDialogFunc.Invoke(0);
        }

        #endregion Command Methods

    }
}