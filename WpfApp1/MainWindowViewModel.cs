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
    public class MainWindowViewModel : ViewModelBase
    {

        #region Properties

        private ExcelA _addExcelA;
        private ExcelA _selectedItemExcelA;
        private ExcelB _addExcelB;
        private ExcelB _currExcelBEntry;

        private ObservableCollection<ExcelA> _myCollectionA;
        private ObservableCollection<ExcelA> _myCollectionBForTable;
        private ObservableCollection<ExcelB> _myCollectionB;
        private ObservableCollection<ExcelC> _myCollectionC;
        private ObservableCollection<String> _myCollectionBMains;
        private List<String> _monthComboboxSource;

        private String _myActiveItemB;
        private String _selectedMonth;

        public ExcelA AddExcelA
        {
            get { return _addExcelA; }
            set { SetProperty(ref _addExcelA, value, () => AddExcelA); }
        }

        public ExcelA SelectedItemExcelA
        {
            get { return _selectedItemExcelA; }
            set { SetProperty(ref _selectedItemExcelA, value, () => SelectedItemExcelA); }
        }

        public ExcelB AddExcelB
        {
            get { return _addExcelB; }
            set { SetProperty(ref _addExcelB, value, () => AddExcelB); }
        }

        public ExcelB CurrExcelBEntry
        {
            get { return _currExcelBEntry; }
            set { SetProperty(ref _currExcelBEntry, value, () => CurrExcelBEntry); }
        }

        public ObservableCollection<ExcelA> MyCollectionA
        {
            get { return _myCollectionA; }
            set { SetProperty(ref _myCollectionA, value, () => MyCollectionA); }
        }

        public ObservableCollection<ExcelA> MyCollectionBForTable
        {
            get { return _myCollectionBForTable; }
            set { SetProperty(ref _myCollectionBForTable, value, () => MyCollectionBForTable); }
        }

        public ObservableCollection<ExcelB> MyCollectionB
        {
            get { return _myCollectionB; }
            set { SetProperty(ref _myCollectionB, value, () => MyCollectionB); }
        }

        public ObservableCollection<ExcelC> MyCollectionC
        {
            get { return _myCollectionC; }
            set { SetProperty(ref _myCollectionC, value, () => MyCollectionC); }
        }

        public ObservableCollection<String> MyCollectionBMains
        {
            get { return _myCollectionBMains; }
            set { SetProperty(ref _myCollectionBMains, value, () => MyCollectionBMains); }
        }

        public List<String> MonthComboboxSource
        {
            get { return _monthComboboxSource; }
            set { SetProperty(ref _monthComboboxSource, value, () => MonthComboboxSource); }
        }

        public String MyActiveItemB
        {
            get { return _myActiveItemB; }
            set { SetProperty(ref _myActiveItemB, value, () => MyActiveItemB); OnBItemChangedReloadGrid(value); }
        }

        public String SelectedMonth
        {
            get { return _selectedMonth; }
            set { _LoadAllGridData(value); _firstTimeMonthChanged = false; SetProperty(ref _selectedMonth, value, () => SelectedMonth); }
        }

        #endregion Properties

        #region Data

        private Boolean _firstTimeMonthChanged = true;
        private ExcelManager _manager;

        #endregion Data

        #region Commands&Services

        public Func<int, int, string> ShowAddEntryDialogFuncExcelC { get; set; }

        public ICommand AddEntryCommandExcelA { get; set; }
        public ICommand AddEntryCommandExcelB { get; set; }
        public ICommand AddEntryCommandExcelC { get; set; }
        public ICommand ExportExcelACommand { get; set; }
        public ICommand ExportExcelBCommand { get; set; }
        public ICommand ExportExcelCCommand { get; set; }
        public ICommand DeleteEntryAOnKeyCommand { get; set; }

        public IMessageBoxService MessageBoxService { get { return ServiceContainer.GetService<IMessageBoxService>(); } set { } }
        public IServiceContainer IServiceContainer { get; set; }

        #endregion Commands&Services

        #region Initialisations
        public MainWindowViewModel()
        {
            InitCommandsAndServices();
            InitValuesAndCollections();
        }

        private void InitCommandsAndServices()
        {
            AddEntryCommandExcelA = new DelegateCommand(AddEntryExcelA);
            AddEntryCommandExcelB = new DelegateCommand(AddEntryExcelB);
            AddEntryCommandExcelC = new DelegateCommand(AddEntryExcelC);
            ExportExcelACommand = new DelegateCommand(ExportExcelA);
            ExportExcelBCommand = new DelegateCommand(ExportExcelB);
            ExportExcelCCommand = new DelegateCommand(ExportExcelC);
            DeleteEntryAOnKeyCommand = new DelegateCommand(DeleteEntryAOnKey);

            IServiceContainer = new ServiceContainer(this);
            ServiceContainer.RegisterService(new DevExpress.Mvvm.UI.MessageBoxService());
        }

        private void InitValuesAndCollections()
        {
            // IMPOTANT: The order of these is essential as the last attributes are calling command on set()
            _manager = new ExcelManager();

            AddExcelA = new ExcelA();
            AddExcelB = new ExcelB();

            MyCollectionA = new ObservableCollection<ExcelA>();
            MyCollectionB = new ObservableCollection<ExcelB>();
            MyCollectionC = new ObservableCollection<ExcelC>();

            MonthComboboxSource = Constants.Months;
            SelectedMonth = Constants.getCurrentDateMonth();
        }

        #endregion Initialisations

        #region Public Methods

        public void StoreAllCollections()
        {
            this._manager.StoreCollectionA(MyCollectionA.ToList(), SelectedMonth);
            this._manager.StoreCollectionB(MyCollectionB.ToList());
            this._manager.StoreCollectionC(MyCollectionC.ToList(), SelectedMonth);
        }

        public void AddEntryExcelCFromCodeBehind(ExcelC entry)
        {
            MyCollectionC.Add(entry);
        }

        #endregion Public Methods

        #region Private Methods    
        
        // Load-And-Store-Methods

        private void _LoadAllGridData(String month)

        {
            if (!_firstTimeMonthChanged)
            {
                StoreAllCollections();
            }

            try
            {
                MyCollectionA = new ObservableCollection<ExcelA>(_manager.LoadCollectionExcelA(month));
                MyCollectionB = new ObservableCollection<ExcelB>(_manager.LoadCollectionExcelB());
                MyCollectionC = new ObservableCollection<ExcelC>(_manager.LoadCollectionExcelC(month));
            }
            catch (System.IO.IOException e)
            {
                var x = MessageBoxService.ShowMessage("Close all the opened excel files and restart the app!", "Error", MessageButton.OK, MessageIcon.Information);
                if (x.Equals(MessageResult.OK))
                    Environment.Exit(0);
                return;
            }

            MyCollectionBMains = new ObservableCollection<String>(_manager.LoadCollectionBMains(MyCollectionB.ToList()));
            MyCollectionBForTable = new ObservableCollection<ExcelA>();
        }

        private void OnBItemChangedReloadGrid(String newValue)
        {
            MyCollectionBForTable = new ObservableCollection<ExcelA>(_manager.LoadCollectionForExcelBValue(newValue, MyCollectionA.ToList()));
            foreach (var x in MyCollectionB)
            {
                if (x.Val3Main.Equals(newValue))
                {
                    CurrExcelBEntry = x;
                    break;
                }
            }
        }

        #endregion Private Methods

        #region Command Methods

        // CRUD-Commands

        private void AddEntryExcelA()
        {           
            MyCollectionA.Add(AddExcelA);    
            OnBItemChangedReloadGrid(MyActiveItemB);
            AddExcelA = new ExcelA();
        }

        private void AddEntryExcelB()
        {
            MyCollectionB.Add(AddExcelB);
            MyCollectionBMains = new ObservableCollection<String>(_manager.LoadCollectionBMains(MyCollectionB.ToList()));
            _manager.StoreCollectionB(MyCollectionB.ToList());
            AddExcelB = new ExcelB();
        }

        private void AddEntryExcelC()
        {
            ShowAddEntryDialogFuncExcelC.Invoke(0, 0);
        }

        private void DeleteEntryAOnKey()
        {
            if (MyCollectionA.Contains(SelectedItemExcelA))
            {
                MyCollectionA.Remove(SelectedItemExcelA);
                OnBItemChangedReloadGrid(MyActiveItemB);
            }
        }

        // Export-Commands

        private void ExportExcelA()
        {
            _manager.StoreCollectionA(MyCollectionA.ToList(), SelectedMonth);
            _manager.OpenExcelA(SelectedMonth);
        }

        private void ExportExcelB()
        {
            _manager.StoreCollectionB(MyCollectionB.ToList());
            _manager.OpenExcelB();
        }

        private void ExportExcelC()
        {
            _manager.StoreCollectionC(MyCollectionC.ToList(), SelectedMonth);
            _manager.OpenExcelC(SelectedMonth);
        }

        #endregion Command Methods

    }

}
