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
    //ctrl + m + o: collapse all
    //ctrl + m + l: expand all
{
    public class MainWindowViewModel : ViewModelBase
    {

        #region Properties

        private ObservableCollection<String> _MyCollectionExcelAValue12;

        public ObservableCollection<String> MyCollectionExcelAValue12
        {
            get { return _MyCollectionExcelAValue12; }
            set { SetProperty(ref _MyCollectionExcelAValue12, value, () => MyCollectionExcelAValue12); }
        }

        private ExcelA _addExcelA;
        private ExcelB _addExcelB;
        private ExcelB _currExcelBEntry;

        private ExcelC _SelectedItemExcelC;

        public ExcelC SelectedItemExcelC
        {
            get { return _SelectedItemExcelC; }
            set { SetProperty(ref _SelectedItemExcelC, value, () => SelectedItemExcelC); }
        }

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
            set { SetProperty(ref _myCollectionA, value, () => MyCollectionA); OnAItemChangedReloadVal4s(); }
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
        public Func<ExcelC, string> ShowEditEntryDialogFuncExcelC { get; set; }

        public ICommand AddEntryCommandExcelA { get; set; }
        public ICommand AddEntryCommandExcelB { get; set; }
        public ICommand AddEntryCommandExcelC { get; set; }
        public ICommand EditEntryCommandExcelC { get; set; }
        public ICommand DeleteEntryCommandExcelC { get; set; }
        public ICommand ExportExcelACommand { get; set; }
        public ICommand ExportExcelBCommand { get; set; }
        public ICommand ExportExcelCCommand { get; set; }
        public ICommand OnGridEditedCommand { get; set; }
        public ICommand SaveExcelsCommand { get; set; }

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
            EditEntryCommandExcelC = new DelegateCommand(EditEntryExcelC);
            DeleteEntryCommandExcelC = new DelegateCommand(DeleteEntryExcelC);
            ExportExcelACommand = new DelegateCommand(ExportExcelA);
            ExportExcelBCommand = new DelegateCommand(ExportExcelB);
            ExportExcelCCommand = new DelegateCommand(ExportExcelC);
            OnGridEditedCommand = new DelegateCommand(OnAItemChangedReloadVal4s);
            SaveExcelsCommand = new DelegateCommand(SaveExcels);

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
            MyCollectionExcelAValue12 = new ObservableCollection<String>();
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

        public void EditEntryExcelCFromCodeBehind(ExcelC entry)
        {
            if(entry != null)
            {
                List<ExcelC> tempList = MyCollectionC.ToList();
                foreach(ExcelC item in tempList)
                {
                    if (item.Val1.Equals(entry.Val1))
                    {
                        item.Val2 = entry.Val2;
                        item.Val3 = entry.Val3;
                        item.Val4 = entry.Val4;
                        MyCollectionC = new ObservableCollection<ExcelC>(tempList);
                        break;
                    }
                }
            }
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

        private void OnAItemChangedReloadVal4s()
        {
            List<String> tempListStringVal4s = new List<String>();
            foreach(ExcelA item in MyCollectionA.ToList())
            {
                String temp = item.Val1.ToString() + " " + item.Val2.ToString();
                item.Val4 = temp;
                tempListStringVal4s.Add(temp);
            }
            MyCollectionExcelAValue12 = new ObservableCollection<String>(tempListStringVal4s);
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

        private void EditEntryExcelC()
        {
            ShowEditEntryDialogFuncExcelC.Invoke(SelectedItemExcelC);
        }

        private void DeleteEntryExcelC()
        {
            if(SelectedItemExcelC != null && SelectedItemExcelC.Val1 != null)
            {
                MyCollectionC.Remove(SelectedItemExcelC);
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

        private void SaveExcels()
        {
            StoreAllCollections();
        }

        #endregion Command Methods

    }

}
