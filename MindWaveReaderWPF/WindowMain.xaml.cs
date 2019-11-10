using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.IO;
using NeuroSky.ThinkGear;
using NeuroSky.ThinkGear.Algorithms;
using System.Linq;
using System.Collections;
using LiveCharts;
using LiveCharts.Wpf;
using System.Windows.Media;
using System.Text;
using System.Windows.Forms;

namespace MindWaveReaderWPF
{
    /// <summary>
    /// Interaction logic for WindowMain.xaml
    /// </summary>
    public partial class WindowMain : Window
    {
        private readonly int _RESET_WAITING_STEPS = 1000;
        public static int chksq = 1;

        public WindowMain()
        {
            InitializeComponent();

            this.Closing += WindowMain_Closing;

        }
        private void DoEvent()
        {
            try
            {
                System.Windows.Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new Action(delegate { }));
                this.UpdateLayout();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        public void LogAdd(string logLine)
        {
            try
            {
                this.Dispatcher.Invoke(() =>
                {
                    if (this.TextBlockLog?.Text?.Split(new char[] { '\n' })?.Length > 1000)
                    {
                        this.TextBlockLog.Text = string.Empty;
                    }
                    this.TextBlockLog.Text += logLine + "\r\n";
                    this.TextBlockLog.InvalidateVisual();
                }, DispatcherPriority.ContextIdle);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        public void LogClear()
        {
            try
            {
                this.Dispatcher.Invoke(() =>
                {
                    this.TextBlockLog.Text = null;
                    this.TextBlockLog.InvalidateVisual();
                }, DispatcherPriority.ContextIdle);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private void ButtonConnect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //this.NativeConnect();
                this.TGConnect();

                // UI Update
                this.ButtonConnect.IsEnabled = false;
                this.ButtonDisconnect.IsEnabled = true;
            }
            catch (Exception ex)
            {
                this.LogAdd("Exception: ButtonConnect_Click");
                this.LogAdd(ex.Message);
            }
        }
        private void ButtonDisconnect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //this.NativeDisconnect();
                this.TGDisconnect();

                // UI Update
                this.ButtonConnect.IsEnabled = true;
                this.ButtonDisconnect.IsEnabled = false;
            }
            catch (Exception ex)
            {
                this.LogAdd("Exception: ButtonDisconnect_Click");
                this.LogAdd(ex.Message);
            }
        }
        private async void ButtonDemo_Click(object sender, RoutedEventArgs e)
        {
            int time = 23000;  // Simulation times *1000 + @ = 20번 * 1000 + 3000 = 23000
            try
            {
                SQZero.Visibility = Visibility.Hidden;
                CSVexport.Visibility = Visibility.Hidden;
                ButtonDemo.Visibility = Visibility.Hidden;
                this.TGMockDataGenerator();
                await Task.Delay(time);
                GaugeSignal.Value = 0;
                GaugeAttention.Value = 0;
                GaugeMeditation.Value = 0;
                GaugeBlinkStrength.Value = 0;
                GaugeMentalEffort.Value = 0;
                GaugeTaskFamiliarity.Value = 0;
                chksq = 0;
                System.Windows.MessageBox.Show("시뮬레이션이 완료 되었습니다. 잠시만 기다려 주세요..", "Simulation Finished!", MessageBoxButton.OK, MessageBoxImage.Warning);
                await Task.Delay(5000);
                CSVexport.Visibility = Visibility.Visible;
                ButtonDemo.Visibility = Visibility.Visible;

            }
            catch (Exception ex)
            {
                this.LogAdd("Exception: ButtonDemo_Click");
                this.LogAdd(ex.Message);
            }
        }


        private void WindowMain_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //this.NativeDisconnect();
            this.TGDisconnect();
        }


        #region Mind Developer Tools (MDT) - ThinkGear SDK for .NET (Platform x86)
        private Connector TGConnector = null;
        private LimitedQueue<TGData> TGQueue = new LimitedQueue<TGData>(10);
        private TGData TGLatestData = new TGData();
        private void TGConnect()
        {
            TGConnector = new Connector();
            TGConnector.DeviceFound += TGConnector_DeviceFound;
            TGConnector.DeviceNotFound += TGConnector_DeviceNotFound;
            TGConnector.DeviceValidating += TGConnector_DeviceValidating;
            TGConnector.DeviceConnected += TGConnector_DeviceConnected;
            TGConnector.DeviceConnectFail += TGConnector_DeviceConnectFail;
            TGConnector.DeviceDisconnected += TGConnector_DeviceDisconnected;
            TGConnector.setBlinkDetectionEnabled(true);

            // Mental Effort (BETA)
            // Mental Effort measures how hard the subject’s brain is working, 
            // i.e.the amount of workload involved in the task. 
            // Mental Effort algorithm can be used for both within-trial monitoring 
            // (continuous real - time tracking) and between - trial comparison.
            // A trial can be of any length equal to or more than 1 minute.
            // In each trial, the first output index will be given out after the 
            // first minute and new output indexes will then be generated at time 
            // interval defined by the output rate(default: 10s).
            TGConnector.setMentalEffortEnable(true);

            // Familiarity (BETA)
            // Familiarity measures how well the subject is learning a new task or how 
            // well his performance is with certain task. 
            // Familiarity algorithm can be used for both within-trial monitoring 
            // (continuous real-time tracking) and between-trial comparison. 
            // A trial can be of any length equal to or more than 1 minute. In each trial, 
            // the first output index will be given out after the first minute and new output 
            // indexes will then be generated at time interval defined by the output rate (default: 10s).
            TGConnector.setTaskFamiliarityEnable(true);

            // Scan for devices across COM ports
            // The COM port named will be the first COM port that is checked.
            TGConnector.ConnectScan(this.ComboBoxPortName.Text);
        }
        private void TGDisconnect()
        {
            if (TGConnector != null)
            {
                TGConnector.Close();
                TGConnector.Disconnect();
            }
            TGConnector = null;
        }
        private void TGConnector_DeviceFound(object sender, EventArgs e)
        {
            Connector.DeviceEventArgs devArgs = (Connector.DeviceEventArgs)e;
            this.LogAdd("기기를 찾았습니다.");
            this.LogAdd($"Device ID: {devArgs.Device.HeadsetID}");
            this.LogAdd($"연결 포트: { devArgs.Device.PortName}");
        }
        private void TGConnector_DeviceNotFound(object sender, EventArgs e)
        {
            this.LogAdd("기기를 찾지 못하였습니다!");
        }
        private void TGConnector_DeviceValidating(object sender, EventArgs e)
        {
            var ConnectionArgs = ((e != null) ? (Connector.ConnectionEventArgs)e : null);
            this.LogAdd($"{ConnectionArgs?.Connection?.PortName ?? ""} 포트 연결중 ...");
        }
        private void TGConnector_DeviceConnected(object sender, EventArgs e)
        {
            Connector.DeviceEventArgs devArgs = (Connector.DeviceEventArgs)e;
            this.LogAdd("기기 연결 완료");
            if (devArgs?.Device != null)
            {
                this.LogAdd($"헤드셋 ID: {devArgs.Device.HeadsetID}");
                this.LogAdd($"포트: { devArgs.Device.PortName}");
                this.LogAdd($"Data Received Rate: { devArgs.Device.DataReceivedRate}");
                this.LogAdd($"마지막 업데이트: { devArgs.Device.lastUpdate:yyyy/MM/dd hh:mm:ss}");
            }
            devArgs.Device.DataReceived += TGConnector_DataReceived;
        }
        private void TGConnector_DeviceConnectFail(object sender, EventArgs e)
        {
            this.LogAdd($"기기를 찾지 못하였습니다!");
        }
        private void TGConnector_DeviceDisconnected(object sender, EventArgs e)
        {
            this.LogAdd($"기기와 연결을 중단 합니다.");
        }
        private void TGConnector_DataReceived(object sender, EventArgs e)
        {
            Device device = (Device)sender;

            // Parse Data #1
            var TGParser = new TGParser();
            TGParser.Read((e as Device.DataEventArgs).DataRowArray);

            // Parse Data #2
            var TGDataRow = new TGData();
            var TGDataProperties = typeof(TGData).GetProperties();
            for (int i = 0; i < TGParser.ParsedData.Length; i++)
            {
                foreach (var Key in TGParser.ParsedData[i].Keys)
                {
                    if (Key == "Time") continue;
                    if (Key == "Raw") continue;
                    var PropertyInfo = TGDataProperties.Where(prop => prop.Name == Key).FirstOrDefault();
                    if (PropertyInfo == null)
                    {
                        Debug.Assert(false, "정보가 완전히 전달 되지 않았습니다!");
                        continue;
                    }
                    PropertyInfo.SetValue(TGDataRow, TGParser.ParsedData[i][PropertyInfo.Name]);

                    // IMP: This algorithm is resource and computation intensive. 
                    // If you need to run with the Debugger, be aware that this calculation may take many minutes to complete 
                    // when the debugger is engaged. It will output its results only after its calculations are complete.
                    //
                    // If these methods are called before the MSG_MODEL_IDENTIFIED has been received, 
                    // it is considered a request to be processed when the connected equipment is identified. 
                    // It is possible to Enable this feature and later find that it is no longer enabled. 
                    // Once the connected equipment has been identified, if the request is incompatible with the hardware 
                    // or software it will be overridden and the MSG_ERR_CFG_OVERRIDE message sent to provide notification.
                    if (Key == "MSG_MODEL_IDENTIFIED")
                    {
                        LogAdd("MSG_MODEL_IDENTIFIED Recevied...");
                        TGConnector.setMentalEffortRunContinuous(true);
                        TGConnector.setMentalEffortEnable(true);
                        TGConnector.setTaskFamiliarityRunContinuous(true);
                        TGConnector.setTaskFamiliarityEnable(true);
                        TGConnector.setPositivityEnable(false);

                        // the following are included to demonstrate the overide messages
                        TGConnector.setRespirationRateEnable(true); // not allowed with EEG
                        TGConnector.setPositivityEnable(true);// not allowed when famil/diff are enabled
                    }

                    // If these methods are called before the MSG_MODEL_IDENTIFIED has been received, 
                    // it is considered a request to be processed when the connected equipment is identified. 
                    // It is possible to Enable this feature and later find that it is no longer enabled. 
                    // Once the connected equipment has been identified, if the request is incompatible with the hardware 
                    // or software it will be overridden and the MSG_ERR_CFG_OVERRIDE message sent to provide notification.
                    if (TGParser.ParsedData[i].ContainsKey("MSG_ERR_CFG_OVERRIDE"))
                    {
                        LogAdd($"ErrorConfigurationOverride: {TGParser.ParsedData[i]["MSG_ERR_CFG_OVERRIDE"]} Recevied...");
                    }
                    if (TGParser.ParsedData[i].ContainsKey("MSG_ERR_NOT_PROVISIONED"))
                    {
                        LogAdd($"ErrorModuleNotProvisioned: {TGParser.ParsedData[i]["MSG_ERR_NOT_PROVISIONED"]} Recevied...");
                    }
                }
            }

            // Add TGDataRow to Queue
            if (TGDataRow.EegPowerDelta != 0) TGQueue.Enqueue(TGDataRow);
            this.RefreshUI(TGDataRow);
        }
        private async void TGMockDataGenerator()
        {
            int cycle = 1;


            for (cycle = 30; cycle < 51; cycle++)
            {
                var TGDataRow = new TGData();
                var Rand = new Random((int)DateTime.Now.Ticks);
                // Generate Randome Data 30%
                if (Rand.Next(1, 101) < 30) TGDataRow.PoorSignal = Rand.Next(0, 30);
                if (Rand.Next(1, 101) < 30) TGDataRow.Attention = Rand.Next(0, 101);
                if (Rand.Next(1, 101) < 30) TGDataRow.Meditation = Rand.Next(0, 101);

                // Generate Randome Data 10%
                if (Rand.Next(1, 101) < 10) TGDataRow.BlinkStrength = Rand.Next(0, 256);
                if (Rand.Next(1, 101) < 10) TGDataRow.TaskFamiliarity = Rand.Next(-1000, 0);
                if (Rand.Next(1, 101) < 10) TGDataRow.TaskDifficulty = Rand.Next(0, 1000);
                if (Rand.Next(1, 101) < 10) TGDataRow.MentalEffort = Rand.Next(0, 1000);

                // Generate Randome in every row
                TGDataRow.EegPowerDelta = Rand.Next(0, 2000000);
                TGDataRow.EegPowerTheta = Rand.Next(0, 1000000);
                TGDataRow.EegPowerAlpha1 = Rand.Next(0, 200000);
                TGDataRow.EegPowerAlpha2 = Rand.Next(0, 200000);
                TGDataRow.EegPowerBeta1 = Rand.Next(0, 100000);
                TGDataRow.EegPowerBeta2 = Rand.Next(0, 50000);
                TGDataRow.EegPowerGamma1 = Rand.Next(0, 50000);
                TGDataRow.EegPowerGamma2 = Rand.Next(0, 100000);
                // Add TGDataRow to Queue
                TGQueue.Enqueue(TGDataRow);
                this.RefreshUI(TGDataRow);
                await Task.Delay(1000);
            }
        }
        /// <summary>
        /// Refreshes the UI.
        /// </summary>
        /// <param name="currentTGData">The current tg data.</param>
        private void RefreshUI(TGData currentTGData)
        {
            this.Dispatcher.Invoke(() =>
            {
                try
                {
                    // Gauge Signal Quality
                    if (currentTGData.PoorSignal != -1)
                    {
                        var SignalQuality = Math.Round((200D - currentTGData.PoorSignal) / 2D, 1);
                        if (SignalQuality < 0) SignalQuality = 0;
                        GaugeSignal.Value = SignalQuality;
                        TGLatestData.PoorSignal = currentTGData.PoorSignal;
                    }

                    // Gauge ATTENTION/MEDITATION
                    if (currentTGData.Attention != -1) { GaugeAttention.Value = TGLatestData.Attention = currentTGData.Attention; }
                    if (currentTGData.Meditation != -1) { GaugeMeditation.Value = TGLatestData.Meditation = currentTGData.Meditation; }

                    // AngularGauges: BlinkStrength
                    if (currentTGData.BlinkStrength != -1)
                    {
                        GaugeBlinkStrength.Value = (currentTGData.BlinkStrength / 255D * 100D);
                        GaugeBlinkStrength.Tag = _RESET_WAITING_STEPS;
                        TGLatestData.BlinkStrength = currentTGData.BlinkStrength;
                    }
                    if (GaugeBlinkStrength.Tag != null)
                    {
                        var CleanTime = (int)GaugeBlinkStrength.Tag - 1;
                        if (CleanTime == 0)
                        {
                            GaugeBlinkStrength.Tag = null;
                            GaugeBlinkStrength.Value = 0;
                            GaugeBlinkStrength.NeedleFill = Brushes.Black;
                        }
                        else
                        {
                            GaugeBlinkStrength.Tag = CleanTime;
                            GaugeBlinkStrength.NeedleFill = Brushes.Gold;
                        }
                    }

                    // AngularGauges: TaskFamiliarity
                    if (currentTGData.TaskFamiliarity != 0)
                    {
                        GaugeTaskFamiliarity.Value = currentTGData.TaskFamiliarity;
                        GaugeTaskFamiliarity.Tag = _RESET_WAITING_STEPS;
                        TGLatestData.TaskFamiliarity = currentTGData.TaskFamiliarity;
                    }
                    if (GaugeTaskFamiliarity.Tag != null)
                    {
                        var CleanTime = (int)GaugeTaskFamiliarity.Tag - 1;
                        if (CleanTime == 0)
                        {
                            GaugeTaskFamiliarity.Tag = null;
                            GaugeTaskFamiliarity.Value = 0;
                            GaugeTaskFamiliarity.NeedleFill = Brushes.Black;
                        }
                        else
                        {
                            GaugeTaskFamiliarity.Tag = CleanTime;
                            GaugeTaskFamiliarity.NeedleFill = Brushes.Gold;
                        }
                    }

                    // AngularGauges: MentalEffort
                    if (currentTGData.MentalEffort != 0)
                    {
                        GaugeMentalEffort.Value = currentTGData.MentalEffort;
                        GaugeMentalEffort.Tag = _RESET_WAITING_STEPS;
                        TGLatestData.MentalEffort = currentTGData.MentalEffort;
                    }
                    if (GaugeMentalEffort.Tag != null)
                    {
                        var CleanTime = (int)GaugeMentalEffort.Tag - 1;
                        if (CleanTime == 0)
                        {
                            GaugeMentalEffort.Tag = null;
                            GaugeMentalEffort.Value = 0;
                            GaugeMentalEffort.NeedleFill = Brushes.Black;
                        }
                        else
                        {
                            GaugeMentalEffort.Tag = CleanTime;
                            GaugeMentalEffort.NeedleFill = Brushes.Gold;
                        }
                    }

                    // Chart Data / EEG Data
                    if (currentTGData.EegPowerDelta != 0)
                    {
                        var ListLineSeries = new List<LineSeries>();
                        ListLineSeries.Add(new LineSeries() { Title = "Delta Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerDelta).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "Theta Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerTheta).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "Low Alpha Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerAlpha1).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "High Alpha Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerAlpha2).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "Low Beta Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerBeta1).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "High Beta Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerBeta2).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "Low Gamma Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerGamma1).ToArray()), PointGeometry = null, });
                        ListLineSeries.Add(new LineSeries() { Title = "High Gamma Power", Values = new ChartValues<double>(TGQueue.Select(q => q.EegPowerGamma2).ToArray()), PointGeometry = null, });
                        var TGChartSeriesCollection = new SeriesCollection();
                        TGChartSeriesCollection.AddRange(ListLineSeries);
                        this.CartesianChartWaves.AxisX[0].Labels = TGQueue.Select(q => q.TimeStamp.ToString()).ToArray();
                        this.CartesianChartWaves.AxisY[0].LabelFormatter = value => value.ToString("N0");
                        this.CartesianChartWaves.Series = TGChartSeriesCollection;

                        TGLatestData.EegPowerDelta = currentTGData.EegPowerDelta;
                        TGLatestData.EegPowerTheta = currentTGData.EegPowerTheta;
                        TGLatestData.EegPowerAlpha1 = currentTGData.EegPowerAlpha1;
                        TGLatestData.EegPowerAlpha2 = currentTGData.EegPowerAlpha2;
                        TGLatestData.EegPowerBeta1 = currentTGData.EegPowerBeta1;
                        TGLatestData.EegPowerBeta2 = currentTGData.EegPowerBeta2;
                        TGLatestData.EegPowerGamma1 = currentTGData.EegPowerGamma1;
                        TGLatestData.EegPowerGamma2 = currentTGData.EegPowerGamma2;
                    }

                    if (currentTGData.TaskDifficulty != 0) TGLatestData.TaskDifficulty = currentTGData.TaskDifficulty;
                    if (currentTGData.RawCh1 != 0) TGLatestData.RawCh1 = currentTGData.RawCh1;
                    if (currentTGData.RawCh2 != 0) TGLatestData.RawCh2 = currentTGData.RawCh2;
                    if (currentTGData.RawCh3 != 0) TGLatestData.RawCh3 = currentTGData.RawCh3;
                    if (currentTGData.RawCh4 != 0) TGLatestData.RawCh4 = currentTGData.RawCh4;
                    if (currentTGData.RawCh5 != 0) TGLatestData.RawCh5 = currentTGData.RawCh5;
                    if (currentTGData.RawCh6 != 0) TGLatestData.RawCh6 = currentTGData.RawCh6;
                    if (currentTGData.RawCh7 != 0) TGLatestData.RawCh7 = currentTGData.RawCh7;
                    if (currentTGData.RawCh8 != 0) TGLatestData.RawCh8 = currentTGData.RawCh8;
                    if (currentTGData.MSG_ERR_CFG_OVERRIDE != 0) TGLatestData.MSG_ERR_CFG_OVERRIDE = currentTGData.MSG_ERR_CFG_OVERRIDE;
                    if (currentTGData.MSG_ERR_NOT_PROVISIONED != 0) TGLatestData.MSG_ERR_NOT_PROVISIONED = currentTGData.MSG_ERR_NOT_PROVISIONED;
                    if (currentTGData.MSG_MODEL_IDENTIFIED != 0) TGLatestData.MSG_MODEL_IDENTIFIED = currentTGData.MSG_MODEL_IDENTIFIED;
                    if (currentTGData.RespiratoryRate != 0) TGLatestData.RespiratoryRate = currentTGData.TaskDifficulty;
                    if (currentTGData.Positivity) TGLatestData.Positivity = currentTGData.Positivity;

                    // Update TextLog | نمایش متنی
                    var TextLog = new StringBuilder();
                    TextLog.AppendLine("방금 업데이트 된 값");
                    typeof(TGData).GetProperties().ToList().ForEach(p => TextLog.AppendLine($"{p.Name}: {(p.GetValue(this.TGLatestData)?.ToString() ?? "null")}"));
                    this.TextBlockLog.Text = TextLog.ToString();
                }
                catch (Exception ex)
                {
                    this.LogAdd("Exception: RefreshUI");
                    this.LogAdd(ex.Message);
                }

            }, DispatcherPriority.ContextIdle);
        }
        #endregion

        #region NativeThinkgear Solution (Platform x64)
        public NativeThinkgear Thinkgear = null;
        public int ThinkgearConnectionId = -1;
        private void NativeConnect()
        {
            try
            {
                int ResultCode = -1;
                this.LogClear();
                this.Thinkgear = new NativeThinkgear();

                // Print driver version number
                LogAdd($"TG_GetVersion: {NativeThinkgear.TG_GetVersion()}");

                // Get a connection ID handle to ThinkGear
                ThinkgearConnectionId = NativeThinkgear.TG_GetNewConnectionId();
                LogAdd($"TG_GetNewConnectionId: {ThinkgearConnectionId}");
                if (ThinkgearConnectionId < 0) return;

                // Set/open stream (raw bytes) log file for connection
                //ResultCode = NativeThinkgear.TG_SetStreamLog(ThinkgearConnectionId, "streamLog.txt");
                //LogAdd($"TG_SetStreamLog ResultCode: {ResultCode}");
                //if (ResultCode < 0) return;

                // Set/open data (ThinkGear values) log file for connection
                //ResultCode = NativeThinkgear.TG_SetDataLog(ThinkgearConnectionId, "dataLog.txt");
                //LogAdd($"TG_SetDataLog ResultCode: {ResultCode}");
                //if (ResultCode < 0) return;

                /* Attempt to connect the connection ID handle to serial port "COM3" */
                ResultCode = NativeThinkgear.TG_Connect(
                    ThinkgearConnectionId,
                    $"\\\\.\\{this.ComboBoxPortName.Text}",
                    (NativeThinkgear.Baudrate)Convert.ToInt32(this.ComboBoxBaudRate.Text),
                    NativeThinkgear.SerialDataFormat.TG_STREAM_PACKETS);
                LogAdd($"TG_Connect ResultCode: {ResultCode}");
                if (ResultCode < 0) return;
            }
            catch (Exception ex)
            {
                this.LogAdd("Exception: ButtonConnect_Click");
                this.LogAdd(ex.Message);
            }

            // Start Data Reader
            this.NativeProcessesDoWork();
        }
        private void NativeDisconnect()
        {
            try
            {
                this.LogClear();

                // Reset Data
                ThinkgearConnectionId = -1;
                this.Thinkgear = null;

                // Stop Data Reader
                Thread.Sleep(1000);
                //this.backgroundWorkerProcesses_DoWork();

                // Disconnect test
                NativeThinkgear.TG_Disconnect(ThinkgearConnectionId);
                LogAdd($"TG_Disconnect");

                // Clean up
                NativeThinkgear.TG_FreeConnection(ThinkgearConnectionId);
                LogAdd($"TG_FreeConnection");
            }
            catch (Exception ex)
            {
                this.LogAdd("Exception: ButtonDisconnect_Click");
                this.LogAdd(ex.Message);
            }
        }
        private void NativeProcessesDoWork()
        {
            try
            {
                while (ThinkgearConnectionId != -1 && this.Thinkgear != null)
                {
                    this.DoEvent();

                    int ResultCode = NativeThinkgear.TG_ReadPackets(ThinkgearConnectionId, -1);
                    if (ResultCode != 1) continue;

                    bool hasValue = false;
                    var Values = new Dictionary<NativeThinkgear.DataType, float>();
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_POOR_SIGNAL, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_ATTENTION, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_MEDITATION, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_ALPHA1, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_ALPHA2, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_BETA1, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_BETA2, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_DELTA, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_THETA, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_GAMMA1, ref Values);
                    hasValue |= NativeGetData(NativeThinkgear.DataType.TG_DATA_GAMMA2, ref Values);


                    if (hasValue)
                    {
                        this.LogClear();
                        this.LogAdd($"{DateTime.Now:yyyy/MM/dd hh:mm:ss}");
                        foreach (NativeThinkgear.DataType dataType in Enum.GetValues(typeof(NativeThinkgear.DataType)))
                        {
                            if (Values.ContainsKey(dataType))
                            {
                                this.LogAdd($"{dataType.ToString()}: {Values[dataType]}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogAdd("Exception: ProcessesDoWork");
                this.LogAdd(ex.Message);
            }
        }
        private bool NativeGetData(NativeThinkgear.DataType dataType, ref Dictionary<NativeThinkgear.DataType, float> values)
        {
            if (ThinkgearConnectionId < 0) return false;
            if (NativeThinkgear.TG_GetValueStatus(ThinkgearConnectionId, dataType) != 0)
            {
                values.Add(dataType, NativeThinkgear.TG_GetValue(ThinkgearConnectionId, dataType));
                return true;
            }
            return false;
        }
        #endregion

        #region Saving_Result_As_CSV
        public async void CSVexport_Click(object sender, RoutedEventArgs e)
        {
            InitializeComponent();
            chksq = 1;
            try
            {
                SaveFileDialog saveAsfile = new SaveFileDialog();
                saveAsfile.InitialDirectory = @"C:\";
                saveAsfile.Title = "다른 이름으로 저장";
                saveAsfile.Filter = "CSV Document(*.csv)|*.csv|All Files(*.*)|*.*";
                saveAsfile.DefaultExt = "csv";
                saveAsfile.AddExtension = true;
                if (saveAsfile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (TGLatestData.PoorSignal != -1)
                    {
                        var SignalQuality = Math.Round((200D - TGLatestData.PoorSignal) / 2D, 1);
                        if (SignalQuality < 0) SignalQuality = 0;
                        GaugeSignal.Value = SignalQuality;
                        TGLatestData.PoorSignal = TGLatestData.PoorSignal;
                    }
                    SQZero.Visibility = 0;
                    CSVexport.Visibility = Visibility.Hidden;
                    FileStream filestream = new FileStream(saveAsfile.FileName, FileMode.Create, FileAccess.Write);
                    System.Windows.MessageBox.Show("CSV 파일이 Writing 됩니다...\nWriting 을 종료하시고 싶다면, 측정 종료 버튼을 눌러주세요\n시뮬레이션을 돌리시면, 시뮬레이션이 끝난 후 자동으로 저장 됩니다.,", "CSV Writing...", MessageBoxButton.OK, MessageBoxImage.Warning);
                    System.IO.StreamWriter file = new System.IO.StreamWriter(filestream);
                    int time = 1000;
                    file.WriteLine("Signal Quality, Delta, Theta, Alpha1, Alpha2, Beta1, Beta2, Gamma1, Gamma2, Attention, Meditation, Mental Effort, Task Familiarity, Task Difficulty, Blink Strength");
                    while (true)
                    {
                        if (chksq == 0)
                        {
                            int EndTime = 4000;
                            file.Close();
                            await Task.Delay(EndTime);
                            System.Windows.MessageBox.Show("CSV 저장 완료!", "CSV Saved", MessageBoxButton.OK, MessageBoxImage.Warning);
                            SQZero.Visibility = Visibility.Hidden;
                            CSVexport.Visibility = Visibility.Visible;
                            while (true) await Task.Delay(1000);
                        }
                        await Task.Delay(time);
                        file.WriteLine("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}", GaugeSignal.Value, TGLatestData.EegPowerDelta, TGLatestData.EegPowerTheta, TGLatestData.EegPowerAlpha1, TGLatestData.EegPowerAlpha2, TGLatestData.EegPowerBeta1, TGLatestData.EegPowerBeta2, TGLatestData.EegPowerGamma1, TGLatestData.EegPowerGamma2, TGLatestData.Attention, TGLatestData.Meditation, TGLatestData.MentalEffort, TGLatestData.TaskFamiliarity, TGLatestData.TaskDifficulty, TGLatestData.BlinkStrength);
                    }

                }
            }
            catch (ApplicationException ex)
            {
                throw new ApplicationException("ApplicationException!", ex);
            }
        }
        public void SQzero_Click(object sender, RoutedEventArgs e)
        {
            SQZero.Visibility = Visibility.Hidden;
            System.Windows.MessageBox.Show("CSV Writing이 중단되고 저장중입니다... \n잠시만 기다려 주세요..", "CSV Saving..", MessageBoxButton.OK, MessageBoxImage.Warning);
            chksq = 0;
        }
        #endregion
    }
}
