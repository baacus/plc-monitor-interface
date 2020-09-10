using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace PLC_Monitor_Interface
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<Variable> varList { get; set; }

        BackgroundWorker worker;
        DispatcherTimer donguTimer;
        LibnoDaveClass libno;

        int cycleTime;
        string plcTableName = "plcTable";
        string varTableName = "varTable";
        List<string> plcTableCols = new List<string>() { "IP", "CycleTime" };
        List<string> varTableCols = new List<string>() { "Name", "AType", "VType", 
                                                         "DBNo", "ByteNo", "BitNo", 
                                                         "MonitorValue", "IsModified", "ModifyValue" };
        private PLCStatus _plc;
        private OnlineStatus _ostatus;
        public PLCStatus Plc
        {
            get { return _plc; }
            set 
            { 
                _plc = value;
                if (value == PLCStatus.Disconnected) Ostatus = OnlineStatus.Offline;

                Dispatcher.Invoke(new Action(() =>
                {
                    stack_IP.Background = (value == PLCStatus.Connected) ? Brushes.ForestGreen : Brushes.White;

                    IP_1.IsEnabled = IP_2.IsEnabled =
                    IP_3.IsEnabled = IP_4.IsEnabled = (value == PLCStatus.Connected) ? false : true;

                    btn_Connect.IsEnabled = !(value == PLCStatus.Connected);
                    btn_Disconnect.IsEnabled = (value == PLCStatus.Connected);

                    btn_Online.IsEnabled = (value == PLCStatus.Connected);
                    btn_Offline.IsEnabled = (value == PLCStatus.Connected);
                }));
            }
        }
        public OnlineStatus Ostatus
        {
            get { return _ostatus; }
            set
            {
                _ostatus = value;
                Dispatcher.Invoke(new Action(() =>
                {
                    onlineBar.Visibility = (value == OnlineStatus.Online) ? Visibility.Visible : Visibility.Hidden;
                    btn_Modify.IsEnabled = (value == OnlineStatus.Online);
                    txt_Cycle.IsEnabled = !(value == OnlineStatus.Online);

                }));
            }
        }

        public const double WINDOW_RATIO = 900.0 / 550.0;
        public const double THRES_RATIO = 0.1;
        public enum PLCStatus
        {
            Connected,
            Disconnected
        }
        public enum OnlineStatus
        {
            Online,
            Offline
        }
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            varList = new ObservableCollection<Variable>();
            worker = new BackgroundWorker();
            libno = new LibnoDaveClass();

            Plc = PLCStatus.Disconnected;
            Ostatus = OnlineStatus.Offline;

            donguTimer = new DispatcherTimer() { Interval = new TimeSpan(0,0,0,0,100)};
            donguTimer.Tick += DonguTimer_Tick;
            worker.DoWork += Worker_DoWork;

            donguTimer.Start();
        }

        private void DonguTimer_Tick(object sender, EventArgs e)
        {
            if (!worker.IsBusy)
            {
                worker.RunWorkerAsync();
            }
            cycleTime = txt_Cycle.Text != "" ? int.Parse(txt_Cycle.Text) : 500;
        }
        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            if(Ostatus == OnlineStatus.Online)
            {
                for(int i = 0; i < varList.Count; i++)
                {
                    LibnoDaveClass.AddressType address = (LibnoDaveClass.AddressType)Enum.Parse(typeof(LibnoDaveClass.AddressType), varList[i].AType);
                    VariableType.Variables dataType = (VariableType.Variables) Enum.Parse(typeof(VariableType.Variables), varList[i].VType);
                    int dbNo = varList[i].DBNo;
                    int byteNo = varList[i].ByteNo;
                    int bitNo = varList[i].BitNo;
                    switch (dataType)
                    {
                        case VariableType.Variables.Bit:
                            bool valueBool;
                            libno.read_bit_value(address, dbNo, byteNo, bitNo, out valueBool);
                            varList[i].MonitorValue = valueBool ? "True" : "False";
                            break;
                        case VariableType.Variables.Real:
                            List<float> valueFloat;
                            libno.read_real_values(address, dbNo, byteNo, 1, out valueFloat);
                            varList[i].MonitorValue = valueFloat.Count >= 1 ? valueFloat[0].ToString(CultureInfo.InvariantCulture) : "";
                            break;
                        case VariableType.Variables var when var == VariableType.Variables.Byte || var == VariableType.Variables.UByte :
                            List<int> valueInt;
                            bool signed = var == VariableType.Variables.Byte ? true : false;
                            libno.read_integer_values(address, dbNo, byteNo, 1, out valueInt, LibnoDaveClass.PLCDataType.Byte, signed);
                            varList[i].MonitorValue = valueInt.Count >= 1 ? valueInt[0].ToString() : "";
                            break;
                        case VariableType.Variables var when var == VariableType.Variables.Integer || var == VariableType.Variables.UInteger:
                            List<int> valueInt2;
                            bool signed2 = var == VariableType.Variables.Integer ? true : false;
                            libno.read_integer_values(address, dbNo, byteNo, 1, out valueInt2, LibnoDaveClass.PLCDataType.Integer, signed2);
                            varList[i].MonitorValue = valueInt2.Count >= 1 ? valueInt2[0].ToString() : "";
                            break;
                        case VariableType.Variables var when var == VariableType.Variables.DInteger || var == VariableType.Variables.UDInteger:
                            List<int> valueInt3;
                            bool signed3 = var == VariableType.Variables.DInteger ? true : false;
                            libno.read_integer_values(address, dbNo, byteNo, 1, out valueInt3, LibnoDaveClass.PLCDataType.DInteger, signed3);
                            varList[i].MonitorValue = valueInt3.Count >= 1 ? valueInt3[0].ToString() : "";
                            break;
                    }
                    varList[i].isOnline = true;
                }
                if (!libno.IsConnected)
                    Dispatcher.Invoke(new Action(() => { btn_Disconnect.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); }));
                    
                System.Threading.Thread.Sleep(cycleTime);
            }
            else
            {
                for(int i = 0; i < varList.Count; i++)
                {
                    varList[i].MonitorValue = "";
                    varList[i].isOnline = false;
                }
            }
        }

        private static readonly Regex _regex = new Regex("[^0-9]"); //regex that matches disallowed text
        private static bool IsTextAllowed(string text)
        {
            bool is_match = _regex.IsMatch(text);
            return !is_match;
        }
        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = ! IsTextAllowed(e.Text);
        }
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox obj = (TextBox)sender;
            int lower_int;
            if (int.TryParse(obj.Text, out lower_int))
            {
                if (lower_int > 255)
                {
                    obj.Text = "255";
                }
            }
        }
        private void TextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsTextAllowed(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }
        private void IP_GotFocus(object sender, RoutedEventArgs e)
        {
            ((TextBox)sender).SelectAll();
        }
        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            Window window = (Window)sender;

            if (e.WidthChanged)
            {
                double currentRatio = e.NewSize.Width / e.NewSize.Height;
                if(currentRatio > WINDOW_RATIO + THRES_RATIO || currentRatio < WINDOW_RATIO - THRES_RATIO)
                {
                    window.Height = Math.Round(window.Width / WINDOW_RATIO);
                }
            }
            else if (e.HeightChanged)
            {
                double currentRatio = e.NewSize.Width / e.NewSize.Height;
                if(currentRatio > WINDOW_RATIO + THRES_RATIO || currentRatio < WINDOW_RATIO - THRES_RATIO)
                {
                    window.Width = Math.Round(window.Height * WINDOW_RATIO);
                }
            }
        }


        private void btn_Connect_Click(object sender, RoutedEventArgs e)
        {
            string ip = IP_1.Text + "." + IP_2.Text + "." + IP_3.Text + "." + IP_4.Text;
            libno = new LibnoDaveClass();
            libno.connect_plc(libnoIP: ip);
            if (libno.IsConnected)
            {
                Plc = PLCStatus.Connected;
            }
            else
                MessageBox.Show("Connection could not be established!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        private void btn_Disconnect_Click(object sender, RoutedEventArgs e)
        {
            libno.disconnect_plc();
            if (libno.IsDisconnected)
            {
                Plc = PLCStatus.Disconnected;
            }
        }
        private void btn_Online_Click(object sender, RoutedEventArgs e)
        {
            Ostatus = OnlineStatus.Online;
        }
        private void btn_Offline_Click(object sender, RoutedEventArgs e)
        {
            Ostatus = OnlineStatus.Offline;
        }
        private void btn_Import_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XMLConfig import = new XMLConfig();
                import.create_table(plcTableCols, plcTableName);
                import.create_table(varTableCols, varTableName);
                import.read_config();

                btn_Disconnect.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));

                string[] ip = import.get_columns(plcTableName)["IP"][0].Split('.');
                if (ip.Length >= 4)
                {
                    IP_1.Text = ip[0]; IP_2.Text = ip[1];
                    IP_3.Text = ip[2]; IP_4.Text = ip[3];
                }
                txt_Cycle.Text = import.get_columns(plcTableName)["CycleTime"][0];
                varList.Clear();
                List<List<string>> varConfig = import.get_rows(varTableName);
                for(int i = 0; i < varConfig.Count; i++)
                {
                    Variable variable = new Variable()
                    {
                        Name = varConfig[i][0],
                        AType = varConfig[i][1],
                        VType = varConfig[i][2],
                        DBNo = int.Parse(varConfig[i][3]),
                        ByteNo = int.Parse(varConfig[i][4]),
                        BitNo = int.Parse(varConfig[i][5]),
                        MonitorValue = varConfig[i][6],
                        IsModified = bool.Parse(varConfig[i][7]),
                        ModifyValue = varConfig[i][8]
                    };
                    varList.Add(variable);
                }
            }
            catch
            {
                MessageBox.Show("Import process is not succesful.",
                                 "Error",
                                 MessageBoxButton.OK,
                                 MessageBoxImage.Error);
            }

        }
        private void btn_Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XMLConfig export = new XMLConfig();
                export.create_table(plcTableCols, plcTableName);
                export.create_table(varTableCols, varTableName);
                List<List<string>> plcTableRows = new List<List<string>>
            {
                new List<string>()
                {
                    IP_1.Text + "." + IP_2.Text + "." + IP_3.Text + "." + IP_4.Text,
                    txt_Cycle.Text
                }
            };
                List<List<string>> varTableRows = new List<List<string>>();
                for (int i = 0; i < varList.Count; i++)
                {
                    varTableRows.Add(varList[i].ToList());
                }
                export.add_rows(plcTableRows, plcTableName);
                export.add_rows(varTableRows, varTableName);

                export.save_config();
            }
            catch
            {
                MessageBox.Show("Export process is not succesful.",
                 "Error",
                 MessageBoxButton.OK,
                 MessageBoxImage.Error);
            }

        }
        private void btn_Modify_Click(object sender, RoutedEventArgs e)
        {
            for(int i = 0; i < varList.Count; i++)
            {
                if (varList[i].IsModified)
                {
                    LibnoDaveClass.AddressType address = (LibnoDaveClass.AddressType)Enum.Parse(typeof(LibnoDaveClass.AddressType), varList[i].AType);
                    VariableType.Variables dataType = (VariableType.Variables)Enum.Parse(typeof(VariableType.Variables), varList[i].VType);
                    int dbNo = varList[i].DBNo;
                    int byteNo = varList[i].ByteNo;
                    int bitNo = varList[i].BitNo;
                    switch (dataType)
                    {
                        case VariableType.Variables.Bit:
                            bool valueBool = bool.Parse(varList[i].ModifyValue);
                            libno.write_bit_value(address, dbNo, byteNo, bitNo, valueBool);
                            break;
                        case VariableType.Variables.Real:
                            List<float> valueFloat = new List<float>() { float.Parse(varList[i].ModifyValue,CultureInfo.InvariantCulture) };
                            libno.write_real_values(address, dbNo, byteNo, valueFloat);
                            break;
                        case VariableType.Variables var when var == VariableType.Variables.Byte || var == VariableType.Variables.UByte:
                            List<int> valueInt = new List<int>() { int.Parse(varList[i].ModifyValue) };
                            bool signed = var == VariableType.Variables.Byte ? true : false;
                            libno.write_integer_values(address, dbNo, byteNo, valueInt, LibnoDaveClass.PLCDataType.Byte, signed);;
                            break;
                        case VariableType.Variables var when var == VariableType.Variables.Integer || var == VariableType.Variables.UInteger:
                            List<int> valueInt2 = new List<int>() { int.Parse(varList[i].ModifyValue) };
                            bool signed2 = var == VariableType.Variables.Integer ? true : false;
                            libno.write_integer_values(address, dbNo, byteNo, valueInt2, LibnoDaveClass.PLCDataType.Integer, signed2);
                            break;
                        case VariableType.Variables var when var == VariableType.Variables.DInteger || var == VariableType.Variables.UDInteger:
                            List<int> valueInt3 = new List<int>() { int.Parse(varList[i].ModifyValue) };
                            bool signed3 = var == VariableType.Variables.DInteger ? true : false;
                            libno.write_integer_values(address, dbNo, byteNo, valueInt3, LibnoDaveClass.PLCDataType.DInteger, signed3);
                            break;
                    }
                }
            }
        }

        public class Variable : INotifyPropertyChanged
        {
            private bool _isOnline;
            private string _monitorValue;
            private int _dbNo;
            private int _bitNo;
            private bool _enableModify;
            private string _aType;
            private string _vType;
            private string _modifyValue;
            private bool _isModified { get; set; }
            public string Name { get; set; }
            public int ByteNo { get; set; }
            public int BitNo
            {
                get { return _bitNo; }
                set
                {
                    _bitNo = value > 7 || value < 0 ? 0 : value;
                    OnPropertyChanged(nameof(BitNo));
                }
            }
            public bool isOnline
            {
                get { return _isOnline; }
                set
                {
                    _isOnline = value;
                    OnPropertyChanged(nameof(isOnline));
                }
            }
            public string MonitorValue
            {
                get { return _monitorValue; }
                set
                {
                    _monitorValue = value;
                    OnPropertyChanged(nameof(MonitorValue));
                }
            }
            public int DBNo
            {
                get { return _dbNo; }
                set
                {
                    _dbNo = value < 0 ? 0 : value;
                    OnPropertyChanged(nameof(DBNo));
                }

            }
            public bool IsModified
            {
                get
                { return _isModified; }
                set
                {
                    _isModified = value;
                    OnPropertyChanged(nameof(IsModified));
                }
            }
            public bool EnableModify
            {
                get { return _enableModify; }
                set
                {
                    _enableModify = value;
                    OnPropertyChanged(nameof(EnableModify));
                }
            }
            public string AType
            {
                get { return _aType; }
                set
                {
                    _aType = value;
                    if (value == "Input")
                    {
                        EnableModify = false;
                        IsModified = false;
                    }
                    else
                    {
                        EnableModify = true;
                    }
                    OnPropertyChanged(nameof(AType));
                }
            }
            public string VType
            {
                get { return _vType; }
                set
                {
                    _vType = value;
                    ModifyValue = "0";
                    OnPropertyChanged(nameof(VType));
                }
            }
            public string ModifyValue
            {
                get
                { return _modifyValue; }
                set
                {
                    string result = allowed_entry(value, VType);
                    _modifyValue = result;
                    OnPropertyChanged(nameof(ModifyValue));
                }
            }
            public Variable()
            {
                DBNo = ByteNo = BitNo = 0;
                AType = "DB";
                VType = "Byte";
            }

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }

            public List<string> ToList()
            {
                List<string> lst = new List<string>()
                {
                    Name,
                    AType,
                    VType,
                    DBNo.ToString(),
                    ByteNo.ToString(),
                    BitNo.ToString(),
                    MonitorValue,
                    IsModified.ToString(),
                    ModifyValue
                };
                return lst;
            }

            private string allowed_entry(string value, string variable)
            {
                string result = value;
                Int64 int_result;
                float real_result;
                List<string> int_list = new List<string>() { "Byte", "Integer", "UInteger", "DInteger", "UDInteger" };
                List<Int64> max_list = new List<Int64>() { 255, 32767, 65535, 2147483647, 4294967295 };
                List<Int64> min_list = new List<Int64>() { 0, -32767, 0, -2147483647, 0 };
                if (variable == "Bit")
                {
                    result = (result == "0" || result.ToLower() == "false") ? "False" :
                        ((result == "1" || result.ToLower() == "true") ? "True" : "False");
                }
                else if (variable == "Real")
                {
                    if (!float.TryParse(result, out real_result))
                    {
                        result = "0";
                    }
                }
                else if (int_list.Contains(variable))
                {
                    int index = int_list.IndexOf(variable);
                    if (Int64.TryParse(result, out int_result))
                    {
                        if (int_result > max_list[index]) result = max_list[index].ToString();
                        else if (int_result < min_list[index]) result = min_list[index].ToString();
                    }
                    else result = "0";
                }

                return result;
            }
        }
    }

    public class AddressType : List<string>
    {
        public AddressType()
        {
            this.Add("DB");
            this.Add("Memory");
            this.Add("Input");
            this.Add("Output");
        }
    }
    public class VariableType : List<string>
    {
        public enum Variables
        {
            Bit,
            Byte,
            UByte,
            Integer,
            UInteger,
            DInteger,
            UDInteger,
            Real
        }
        public VariableType()
        {
            List<string> enumList = Enum.GetNames(typeof(Variables)).ToList();
            for(int i = 0; i < enumList.Count; i++)
            {
                this.Add(enumList[i]);
            }
        }
    }

    public class XMLConfig
    {
        public DataSet _xmlDataSet;
        public Dictionary<string,DataTable> _dataTables;

        public XMLConfig()
        {
            _xmlDataSet = new DataSet();
            _dataTables = new Dictionary<string,DataTable>();
        }

        public void create_table(List<string> columnNames, string tableName)
        {
            if (_dataTables.Keys.ToList().Contains(tableName))
            {
                throw new Exception("Table already exists.");
            }

            DataTable table = new DataTable();
            table.TableName = tableName;
            for(int i = 0; i < columnNames.Count; i++)
            {
                table.Columns.Add(columnNames[i]);
            }
            _dataTables.Add(tableName, table);
            _xmlDataSet.Tables.Add(_dataTables[tableName]);

        }

        public Dictionary<string, List<string>> get_columns(string tableName)
        {
            Dictionary<string, List<string>> cols = new Dictionary<string, List<string>>();

            if (!_dataTables.Keys.ToList().Contains(tableName))
            {
                throw new Exception("Table does not exist.");
            }
            DataTable table = _dataTables[tableName];
            for(int i = 0; i < table.Columns.Count; i++)
            {
                cols.Add(table.Columns[i].ColumnName, new List<string>());
                for(int j = 0; j < table.Rows.Count; j++)
                {
                    cols[table.Columns[i].ColumnName].Add((table.Rows[j])[i].ToString());
                }
            }

            return cols;
        }
        public List<List<string>> get_rows(string tableName)
        {
            List<List<string>> rows = new List<List<string>>();
            if (!_dataTables.Keys.ToList().Contains(tableName))
            {
                throw new Exception("Table does not exist.");
            }
            DataTable table = _dataTables[tableName];
            for(int i = 0; i < table.Rows.Count; i++)
            {
                List<string> row = new List<string>();
                for(int j = 0; j < table.Columns.Count; j++)
                {
                    row.Add(table.Rows[i][j].ToString());
                }
                rows.Add(row);
            }
            return rows;
        }
        
        public void add_rows(List<List<string>> rows, string tableName)
        {
            if (!_dataTables.Keys.ToList().Contains(tableName) || rows.Count == 0 ||
               _dataTables[tableName].Columns.Count != rows[0].Count)
            {
                throw new Exception("Table is not implemented.");
            }
            for(int i = 0; i < rows.Count; i++)
            {
                DataRow row = _dataTables[tableName].NewRow();
                for(int j = 0; j < rows[i].Count; j++)
                {
                    row[j] = rows[i][j];
                }
                _dataTables[tableName].Rows.Add(row);
            }
        }

        public void empty_tables()
        {
            foreach(string tableName in _dataTables.Keys.ToList())
            {
                _dataTables[tableName].Rows.Clear();
            }
        }

        public void save_config()
        {
            SaveFileDialog save = new SaveFileDialog();
            save.DefaultExt = ".xml";
            save.Filter = "Xml files (.xml)|*.xml";
            if (save.ShowDialog() == true)
            {
                string path = save.FileName;
                _xmlDataSet.WriteXml(path);
            }
        }

        public void read_config()
        {
            OpenFileDialog open = new OpenFileDialog();
            open.DefaultExt = ".xml";
            open.Filter = "Xml files (.xml)|*.xml";
            if(open.ShowDialog() == true)
            {
                string path = open.FileName;
                try
                {
                    _xmlDataSet.ReadXml(path);
                }
                catch
                {
                    MessageBox.Show("This file does not contain convenient data for the dataset! The file could not be read.", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

    }


}
