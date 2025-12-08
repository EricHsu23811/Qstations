/*****************************************************************************
 **
 **  (c) All Rights Reserved.
 **  
 typedef enum   
{
//--------------- BEGIN CUSTOM AREA ------------------
    stSTATION0,     // Home
    stSTATION1,     // Power Measurement
    stSTATION2,     // CMOS Camera
    stSTATION3,     // EVK-M5
    stSTATION4,     // QA Test Summary
//--------------- END CUSTOM AREA --------------------     
    stINVALID   //=====used for enumerated type range checking (DO NOT REMOVE)=====
}eSTATION;  
 **
 ** Testing Software for Q-Stations
 ** <2025/11/12> based on Thorlabs PM100D Power Meter Driver Example, v_1.0.0.0
 ** <2025/11/19> Add Engineer Mode with password protection
 ** <2025/11/21> Add Version info
 ** <2025/11/24> Add jsonConfigFile read/write for parameter settings storage
 ** <2025/11/26> Add M5 Test Function with M5.Core Python Client in TempTest tabPage, v_1.0.1.0
 ** <2025/12/01> Add M5 Chart Test Function in VCSL tabPage, v_1.0.2.0
 ** <2025/12/03> Add Axis Control Function in Motor Control tabPage, v_1.0.3.0
 **
******************************************************************************/
using M5.Core;  //M5 Python Client
using Newtonsoft.Json;      //
using Newtonsoft.Json.Linq; // For JSON manipulation
using PythonAlgorithmEngine.Core.AlgorithmInput;
using PythonAlgorithmEngine.Core.Config;    //Python Algorithm Engine
using PythonAlgorithmEngine.Core.Engine;
using System;
using System.Diagnostics;
using System.Reflection;    //Assembly.GetExecutingAssembly().GetName()
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Thorlabs.PM100D;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.ComponentModel;


#if X86
using EtherCAT_DLL;
#else
using EtherCAT_DLL_x64;
#endif
using EtherCAT_DLL_Err;

namespace QStations
{
    public partial class Form1 : Form
    {
        /*******************************************************************************
         * Constant             
         ******************************************************************************/
        /* These are constant variables */
        public static readonly string jsonjver = "1.0";
        public static readonly string jsonjver_config = "1.0";
        public static readonly string jsondev = "M5_test";  // 
        public static readonly string userDocFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        /*******************************************************************************
         * Gobal Variable             
         ******************************************************************************/
        private delegate void DelegateShowMessage(string sMessage);
        private delegate void DelegateBoxPrintf(string info, bool clear, Color color);
        static string userAppFolder = userDocFolder + "\\M5test";//@"\StaFlow";        
        string jsonConfigFile = userAppFolder + "\\config.json";
        static string fileName = Assembly.GetExecutingAssembly().GetName().Name;
        string fileVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        string buildDateTime = System.IO.File.GetCreationTime(Assembly.GetExecutingAssembly().Location).ToString("yyyy-MM-dd HH:mm");   //HH:mm:ss

        static DateTime TestDate = DateTime.Now;
        string strFilePath = "";    //userAppFolder + "\\" + "H4_IQC_LBS_BURN.csv";
        static string strDailyLogDate = DateTime.Now.ToString("yyyy-MM-dd");
        static string strFileDailyFolder = userAppFolder + "\\" + DateTime.Now.ToString("yyyy")
            + "\\" + DateTime.Now.ToString("MM") + "\\" + DateTime.Now.ToString("dd");
        static string strLogNameDaily = ""; //strFileDailyFolder + "\\" + "H4_IQC_LBS_BURN" + "_" + strDailyLogDate + ".csv";
        string strFileDailyPath = "";   //userAppFolder + "\\" + strLogNameDaily;    //@"\Burn1to10.csv";

        string strSwVer = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString();  //2023-07-18
        string strFwVer = "";
        string FolderToSaveInAuto;

        string strPassword = "Ab1234";
        private Thorlabs.PM100D.PM100D pm100d;  // power meter object
        string PMresourceName = "USB0::0x1313::0x8078::P0032145::INSTR"; // Power Meter Ressource Name

        //M5 Test code variables
        private readonly M5PythonClient _m5 = new();    //using M5.Core;
        private readonly PythonServerManager _pythonServer = new();

        private readonly AlgorithmEngine _algoEngine;   //using PythonAlgorithmEngine.Core.Engine;

        //Axis control variables
        bool g_bInitialFlag = false;
        ushort g_uRet = 0;
        int g_nSelectMode = 0;
        ushort g_nESCExistCards = 0, g_uESCCardNo = 0;
        ushort[] g_uESCNodeID = new ushort[3];
        ushort[] g_uESCSlotID = new ushort[3];
        ushort[] g_uESCCardNoList = new ushort[32];
        TextBox[] g_pTxtParam = new TextBox[12];
        Label[] g_pLabParamTitle = new Label[12];
        Label[] g_pLabParamUnit = new Label[12];

        public Form1()
        {
            InitializeComponent();

            txtPwdInputPara.PasswordChar = '*'; txtPwdEngCalib.PasswordChar = '*';
            txtPwdEngVcsel.PasswordChar = '*'; txtPwdEngTx.PasswordChar = '*';
            txtPwdEngRx.PasswordChar = '*';

            // 訂閱事件 //M5.core
            _m5.FrameReceived += OnFrameReceived;
            _m5.ErrorOccurred += ex =>
            {
                AppendLog($"[Error] {ex.Message}\r\n");
            };
            _m5.ConnectionStateChanged += state =>
            {
                this.BeginInvoke(new Action(() => UpdateUiByConnectionState(state)));
            };

            string configPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "python_engine_config.json");

            // 初始化 AlgorithmEngine
            var config = PythonEngineConfig.LoadFromFile(configPath);
            _algoEngine = new AlgorithmEngine(config);

            // Axis control variables ------ Set the Interval to 0.1 seconds (100 milliseconds).
            TimCheckStatus.Interval = 100;
            TimCheckStatus.Enabled = true;

            g_pTxtParam[1] = TxtParam01;
            g_pTxtParam[2] = TxtParam02;
            g_pTxtParam[3] = TxtParam03;
            g_pTxtParam[4] = TxtParam04;
            g_pTxtParam[5] = TxtParam05;
            g_pTxtParam[6] = TxtParam06;
            g_pTxtParam[7] = TxtParam07;
            g_pTxtParam[8] = TxtParam08;
            g_pTxtParam[9] = TxtParam09;
            g_pTxtParam[10] = TxtParam10;
            g_pTxtParam[11] = TxtParam11;

            g_pLabParamTitle[1] = LabParam01;
            g_pLabParamTitle[2] = LabParam02;
            g_pLabParamTitle[3] = LabParam03;
            g_pLabParamTitle[4] = LabParam04;
            g_pLabParamTitle[5] = LabParam05;
            g_pLabParamTitle[6] = LabParam06;
            g_pLabParamTitle[7] = LabParam07;
            g_pLabParamTitle[8] = LabParam08;
            g_pLabParamTitle[9] = LabParam09;
            g_pLabParamTitle[10] = LabParam10;
            g_pLabParamTitle[11] = LabParam11;

            g_pLabParamUnit[1] = LabParamUnit01;
            g_pLabParamUnit[2] = LabParamUnit02;
            g_pLabParamUnit[3] = LabParamUnit03;
            g_pLabParamUnit[4] = LabParamUnit04;
            g_pLabParamUnit[5] = LabParamUnit05;
            g_pLabParamUnit[6] = LabParamUnit06;
            g_pLabParamUnit[7] = LabParamUnit07;
            g_pLabParamUnit[8] = LabParamUnit08;
            g_pLabParamUnit[9] = LabParamUnit09;
            g_pLabParamUnit[10] = LabParamUnit10;
            g_pLabParamUnit[11] = LabParamUnit11;

            RdoCSPMode01_CheckedChanged(null, null);

            TrcFeedrate.Maximum = 1000;
            TrcFeedrate.Minimum = 0;
            TrcFeedrate.Value = 100;
            TxtFeedrate.Text = "100";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            grpEngParameter.Enabled = false; grpEngCalibration.Enabled = false;
            grpEngVCSELSetting.Enabled = false; grpTxEngSetting.Enabled = false;
            grpEngRxSetting.Enabled = false;
            radAutoMode.Checked = true;

            lblPMresourceName.Text = PMresourceName; lblSwVer.Text = "Software Version: " + strSwVer;

            //Engineer-Parameter Settings
            trkExpTimeM5Set.Maximum = 10000; trkExpTimeM5Set.Minimum = 0; trkExpTimeM5Set.Value = 1000;
            trkExpTimeM5.Maximum = trkExpTimeM5Set.Maximum; trkExpTimeM5.Minimum = trkExpTimeM5Set.Minimum;
            trkGainM5Set.Maximum = 24; trkGainM5Set.Minimum = 0; trkGainM5Set.Value = 10;
            trkGainM5.Maximum = trkGainM5Set.Maximum; trkGainM5.Minimum = trkGainM5Set.Minimum;
            trkLdPowerM5Set.Maximum = 20; trkLdPowerM5Set.Minimum = 0; trkLdPowerM5Set.Value = 10;
            trkLdPowerM5.Maximum = trkLdPowerM5Set.Maximum; trkLdPowerM5.Minimum = trkLdPowerM5Set.Minimum;

            trkExpTimeCamSet.Maximum = 10000; trkExpTimeCamSet.Minimum = 0; trkExpTimeCamSet.Value = 2000;
            trkExpTimeCam.Maximum = trkExpTimeCamSet.Maximum; trkExpTimeCam.Minimum = trkExpTimeCamSet.Minimum;
            trkGainCamSet.Maximum = 24; trkGainCamSet.Minimum = 0; trkGainCamSet.Value = 5;
            trkGainCam.Maximum = trkGainCamSet.Maximum; trkGainCam.Minimum = trkGainCamSet.Minimum;

            //Engineer-Calibration Settings
            trkGoldPwrSet.Maximum = 100; trkGoldPwrSet.Minimum = 0; trkGoldPwrSet.Value = 75;
            trkGoldPwr.Maximum = trkGoldPwrSet.Maximum; trkGoldPwr.Minimum = trkGoldPwrSet.Minimum;

            toolStripStatusLabel1.Text = "Software Version: " + strSwVer
                /*+ "    固件版本: " + strFwVer*/ + "    built time: " + buildDateTime + "   ";

            tabControl1.TabPages.Remove(tabParam); //tabControl1.TabPages.Remove(tabPage7);   //2025-11-25
            tabControl1.TabPages.Remove(tabEngCalib); tabControl1.TabPages.Remove(tabEngVCSEL);
            tabControl1.TabPages.Remove(tabEngTx); tabControl1.TabPages.Remove(tabEngRx);
            //tabControl1.TabPages.Remove(tabM5Test); tabControl1.TabPages.Remove(tabM5chartTest);
            //tabControl1.TabPages.Remove(tabMotor);
            btnConnect_Click(null, null); //Auto connect M5 on Form Load
            //btnStop_Click(null, null);    //Stop streaming on Form Load
        }

        private void PM100DPwrReadTest_Click(object sender, EventArgs e)   // Measure Power Button
        {
            //double powerValue = MeasurePower();
            //labelPower.Text = (powerValue * 1000).ToString() + " mW";   // Display power in mW
            btnPM100Dtest.Enabled = false;
            try
            {
                pm100d = new Thorlabs.PM100D.PM100D(PMresourceName, false, false);  //  For valid Ressource_Name see NI-Visa documentation.   //2025-12-04 有時會當在這邊
                Thread.Sleep(200); // Wait for 100 ms to impreve stability
                double powerValue;
                int err = pm100d.measPower(out powerValue);
                labelPower.Text = (powerValue * 1000).ToString() + " mW";   // Display power in mW
            }
            catch (BadImageFormatException bie)
            {
                labelPower.Text = bie.Message;
            }
            catch (NullReferenceException nre)
            {
                labelPower.Text = nre.Message;
            }
            catch (ExternalException ex)
            {
                labelPower.Text = ex.Message; //ex = {"Insufficient location information or the device or resource is not present in the system."} //sometimes happens here
                AppendLog($"[PM Error] {ex.Message}\r\n");
                MessageBox.Show("Connect to Power Meter failed.\r\n" + ex.Message);
            }
            finally
            {
                if (pm100d != null)
                    pm100d.Dispose();
            }
            btnPM100Dtest.Enabled = true;
        }
        private void btnPwrValueRead_Click(object sender, EventArgs e)
        {
            double powerValue = MeasurePower();
            labelPower.Text = (powerValue * 1000).ToString() + " mW";   // Display power in mW
        }
        /*
        private double MeasurePower()   // Measure Power Function
        {
            double powerValue = 0;
            try
            {
                pm100d = new Thorlabs.PM100D.PM100D(PMresourceName, false, false);  //  For valid Ressource_Name see NI-Visa documentation.
                int err = pm100d.measPower(out powerValue);
            }
            catch (BadImageFormatException bie)
            {
                AppendLog($"[Error] {bie.Message}\r\n");
                MessageBox.Show(bie.Message);
            }
            catch (NullReferenceException nre)
            {
                AppendLog($"[Error] {nre.Message}\r\n");
                MessageBox.Show(nre.Message);
            }
            catch (ExternalException ex)
            {
                AppendLog($"[Error] {ex.Message}\r\n");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (pm100d != null)
                    pm100d.Dispose();
            }
            toolStripStatusLabel1.Text = (powerValue * 1000).ToString("F6") + " mW";
            return powerValue;
        }
        */

        private double MeasurePower()   // Measure Power Function
        {
            double powerValue = 0;
            try
            {
                pm100d = new Thorlabs.PM100D.PM100D(PMresourceName, false, false); //Hanging here Sometimes  //For valid Ressource_Name see NI-Visa documentation.
                Thread.Sleep(200); // Wait for 100 ms to impreve stability
                int err = pm100d.measPower(out powerValue);
            }
            catch (BadImageFormatException bie)
            {
                AppendLog($"[Error] {bie.Message}\r\n");
                MessageBox.Show(bie.Message);
            }
            catch (NullReferenceException nre)
            {
                AppendLog($"[Error] {nre.Message}\r\n");
                MessageBox.Show(nre.Message);
            }
            catch (ExternalException ex)
            {
                AppendLog($"[Error] {ex.Message}\r\n");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (pm100d != null)
                    pm100d.Dispose();
            }
            toolStripStatusLabel1.Text = (powerValue * 1000).ToString("F6") + " mW";
            
            return powerValue;
        }

        private void cboPmChkResultCalib_TextChanged(object sender, EventArgs e)
        {
            if (cboPmChkResultCalib.Text == "PASS")
            {
                picPmChkResultCalib.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboPmChkResultCalib.ForeColor = Color.Green;
            }
            else if (cboPmChkResultCalib.Text == "fail")
            {
                picPmChkResultCalib.Image = QStations.Properties.Resources.icon_failed_30px;
                cboPmChkResultCalib.ForeColor = Color.Red;
            }
            else
            {
                picPmChkResultCalib.Image = QStations.Properties.Resources.Input_Off_30px;
                cboPmChkResultCalib.ForeColor = Color.Black;
            }
        }
        private void cboCamChkResultCalib_TextChanged(object sender, EventArgs e)
        {
            if (cboCamChkResultCalib.Text == "PASS")
            {
                picCamChkResultCalib.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboCamChkResultCalib.ForeColor = Color.Green;
            }
            else if (cboCamChkResultCalib.Text == "fail")
            {
                picCamChkResultCalib.Image = QStations.Properties.Resources.icon_failed_30px;
                cboCamChkResultCalib.ForeColor = Color.Red;
            }
            else
            {
                picCamChkResultCalib.Image = QStations.Properties.Resources.Input_Off_30px;
                cboCamChkResultCalib.ForeColor = Color.Black;
            }
        }
        private void cboProjectChkResultCalib_TextChanged(object sender, EventArgs e)
        {
            if (cboProjectChkResultCalib.Text == "PASS")
            {
                picProjectChkResultCalib.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboProjectChkResultCalib.ForeColor = Color.Green;
            }
            else if (cboProjectChkResultCalib.Text == "fail")
            {
                picProjectChkResultCalib.Image = QStations.Properties.Resources.icon_failed_30px;
                cboProjectChkResultCalib.ForeColor = Color.Red;
            }
            else
            {
                picProjectChkResultCalib.Image = QStations.Properties.Resources.Input_Off_30px;
                cboProjectChkResultCalib.ForeColor = Color.Black;
            }
        }

        private void btnPwdSend_Click(object sender, EventArgs e)
        {
            string strPassword = "Ab1234";
            if (txtPwdInputPara.Text == strPassword)
            {
                grpEngParameter.Enabled = true;
                txtPwdInputPara.Text = "";
            }
            else if (txtPwdEngCalib.Text == strPassword)
            {
                grpEngCalibration.Enabled = true;
                txtPwdEngCalib.Text = "";
            }
            else if (txtPwdEngVcsel.Text == strPassword)
            {
                grpEngVCSELSetting.Enabled = true;
                txtPwdEngVcsel.Text = "";
            }
            else if (txtPwdEngTx.Text == strPassword)
            {
                grpTxEngSetting.Enabled = true;
                txtPwdEngTx.Text = "";
            }
            else if (txtPwdEngRx.Text == strPassword)
            {
                grpEngRxSetting.Enabled = true;
                txtPwdEngRx.Text = "";
            }
            else
            {
                MessageBox.Show("密碼錯誤，請重新輸入。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                grpEngParameter.Enabled = false;
            }
        }

        private void grpEngineerMode_Leave(object sender, EventArgs e)
        {
            grpEngParameter.Enabled = false; grpEngCalibration.Enabled = false;
            grpEngVCSELSetting.Enabled = false; grpTxEngSetting.Enabled = false; grpEngRxSetting.Enabled = false;
        }

        private void radioButton5_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabEngParam;
        }
        public class MyTransparentPanel : Panel
        {
            public MyTransparentPanel()
            {
                this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
                this.BackColor = Color.Transparent;
            }
        }

        //Engineer-Parameter Settings-------------
        private void chkStreamM5Set_CheckedChanged(object sender, EventArgs e)
        {
            chkStreamM5.Checked = chkStreamM5Set.Checked;
            if (chkStreamM5Set.Checked)
            {
                btnStart_Click(null, null);// Start streaming
            }
            else
            {
                btnStop_Click(null, null); // Stop streaming
            }
        }

        private void cboFpsM5Set_TextChanged(object sender, EventArgs e)
        {
            cboFpsM5.Text = cboFpsM5Set.Text;
        }

        private void chkStreamCamSet_CheckedChanged(object sender, EventArgs e)
        {
            chkStreamCam.Checked = chkStreamCamSet.Checked;
        }

        private void cboFpsCamSet_TextChanged(object sender, EventArgs e)
        {
            cboFpsCam.Text = cboFpsCamSet.Text;
        }

        private void txtExpTimeM5Set_TextChanged(object sender, EventArgs e)
        {
            trkExpTimeM5Set.Value = int.Parse(txtExpTimeM5Set.Text);
            txtExpTimeM5.Text = txtExpTimeM5Set.Text;
            trkExpTimeM5.Value = trkExpTimeM5Set.Value;
        }
        private void trkExpTimeM5Set_ValueChanged(object sender, EventArgs e)
        {
            txtExpTimeM5Set.Text = trkExpTimeM5Set.Value.ToString();
        }


        private void txtExpTimeCamSet_TextChanged(object sender, EventArgs e)
        {
            trkExpTimeCamSet.Value = int.Parse(txtExpTimeCamSet.Text);
            txtExpTimeCam.Text = txtExpTimeCamSet.Text;
            trkExpTimeCam.Value = trkExpTimeCamSet.Value;
        }
        private void trkExpTimeCamSet_ValueChanged(object sender, EventArgs e)
        {
            txtExpTimeCamSet.Text = trkExpTimeCamSet.Value.ToString();
        }

        private void txtGainM5Set_TextChanged(object sender, EventArgs e)
        {
            trkGainM5Set.Value = int.Parse(txtGainM5Set.Text);
            txtGainM5.Text = txtGainM5Set.Text;
            trkGainM5.Value = trkGainM5Set.Value;
        }
        private void trkGainM5Set_ValueChanged(object sender, EventArgs e)
        {
            txtGainM5Set.Text = trkGainM5Set.Value.ToString();
        }
        private void txtGainCamSet_TextChanged(object sender, EventArgs e)
        {
            trkGainCamSet.Value = int.Parse(txtGainCamSet.Text);
            txtGainCam.Text = txtGainCamSet.Text;
            trkGainCam.Value = trkGainCamSet.Value;
        }
        private void trkGainCamSet_ValueChanged(object sender, EventArgs e)
        {
            txtGainCamSet.Text = trkGainCamSet.Value.ToString();
        }

        private void cboLaserEnabledSet_CheckedChanged(object sender, EventArgs e)
        {
            cboLaserEnabled.Checked = cboLaserEnabledSet.Checked;
        }

        private void txtLdPowerM5Set_TextChanged(object sender, EventArgs e)
        {
            trkLdPowerM5Set.Value = int.Parse(txtLdPowerM5Set.Text);
            txtLdPowerM5.Text = txtLdPowerM5Set.Text;
            trkLdPowerM5.Value = trkLdPowerM5Set.Value;
        }

        private void trkLdPowerM5Set_ValueChanged(object sender, EventArgs e)
        {
            txtLdPowerM5Set.Text = trkLdPowerM5Set.Value.ToString();
        }

        private void chkNoiseReductionSet_CheckedChanged(object sender, EventArgs e)
        {
            chkNoiseReduction.Checked = chkNoiseReductionSet.Checked;
        }

        private void cboRoiSet_TextChanged(object sender, EventArgs e)
        {
            cboRoi.Text = cboRoiSet.Text;
        }

        private void cboSavedOptionSet_TextChanged(object sender, EventArgs e)
        {
            cboSavedOption.Text = cboSavedOptionSet.Text;
        }

        //Engineer-Calibration Settings-------------
        private void txtGoldPwrSet_TextChanged(object sender, EventArgs e)
        {
            if (txtGoldPwrSet.Text != "")
            {
                int number;
                if (int.TryParse(txtGoldPwrSet.Text, out number))
                {
                    if (number > 100) { txtGoldPwrSet.Text = "100"; }
                    trkGoldPwrSet.Value = int.Parse(txtGoldPwrSet.Text);
                    txtGoldPwr.Text = txtGoldPwrSet.Text;
                    trkGoldPwr.Value = trkGoldPwrSet.Value;
                }
                else { MessageBox.Show("請輸入數字(0-100)"); txtGoldPwrSet.Text = "0"; }
            }
            else { /*txtGoldPwrSet.Text = "0";*/ }
        }

        private void trkGoldPwrSet_ValueChanged(object sender, EventArgs e)
        {
            txtGoldPwrSet.Text = trkGoldPwrSet.Value.ToString();
        }

        private void txtXHomeEngCalib_TextChanged(object sender, EventArgs e)
        {
            txtXHomeCalib.Text = txtXHomeEngCalib.Text;
        }

        private void txtZHomeEngCalib_TextChanged(object sender, EventArgs e)
        {
            txtZHomeCalib.Text = txtZHomeEngCalib.Text;
        }

        private void radGoldenLaserOnSet_CheckedChanged(object sender, EventArgs e)
        {
            radGoldenLaserOn.Checked = radGoldenLaserOnSet.Checked;
        }

        private void radGoldenLaserOffSet_CheckedChanged(object sender, EventArgs e)
        {
            radGoldenLaserOff.Checked = radGoldenLaserOffSet.Checked;
        }

        private void txtPmPwrEngCalib_TextChanged(object sender, EventArgs e)
        {
            txtPmPwrCalib.Text = txtPmPwrEngCalib.Text;
        }

        private void txtGoldenPwrEngCalib_TextChanged(object sender, EventArgs e)
        {
            txtGoldenPwrCalib.Text = txtGoldenPwrEngCalib.Text;
        }

        private void txtPmLDM1CalibSet_TextChanged(object sender, EventArgs e)
        {
            txtPmLDM1Calib.Text = txtPmLDM1CalibSet.Text;
        }

        private void txtPmLDM2CalibSet_TextChanged(object sender, EventArgs e)
        {
            txtPmLDM2Calib.Text = txtPmLDM2CalibSet.Text;
        }

        private void txtPmSkewCalibSet_TextChanged(object sender, EventArgs e)
        {
            txtPmSkewCalib.Text = txtPmSkewCalibSet.Text;
        }

        private void txtCamLDM1CalibSet_TextChanged(object sender, EventArgs e)
        {
            txtCamLDM1Calib.Text = txtCamLDM1CalibSet.Text;
        }

        private void txtCamLDM2CalibSet_TextChanged(object sender, EventArgs e)
        {
            txtCamLDM2Calib.Text = txtCamLDM2CalibSet.Text;
        }

        private void txtCamSkewCalibSet_TextChanged(object sender, EventArgs e)
        {
            txtCamSkewCalib.Text = txtCamSkewCalibSet.Text;
        }

        private void txtProjectLDM1CalibSet_TextChanged(object sender, EventArgs e)
        {
            txtProjectLDM1Calib.Text = txtProjectLDM1CalibSet.Text;
        }

        private void txtProjectLDM2CalibSet_TextChanged(object sender, EventArgs e)
        {
            txtProjectLDM2Calib.Text = txtProjectLDM2CalibSet.Text;
        }

        private void txtProjectSkewCalibSet_TextChanged(object sender, EventArgs e)
        {
            txtProjectSkewCalib.Text = txtProjectSkewCalibSet.Text;
        }

        private void cboPmChkResultEngCalib_TextChanged(object sender, EventArgs e)
        {
            cboPmChkResultCalib.Text = cboPmChkResultEngCalib.Text;
            if (cboPmChkResultEngCalib.Text == "PASS")
            {
                picPmChkResultEngCalib.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboPmChkResultEngCalib.ForeColor = Color.Green;
            }
            else if (cboPmChkResultEngCalib.Text == "fail")
            {
                picPmChkResultEngCalib.Image = QStations.Properties.Resources.icon_failed_30px;
                cboPmChkResultEngCalib.ForeColor = Color.Red;
            }
            else
            {
                picPmChkResultEngCalib.Image = QStations.Properties.Resources.Input_Off_30px;
                cboPmChkResultEngCalib.ForeColor = Color.Black;
            }
        }

        private void cboCamChkResultEngCalib_TextChanged(object sender, EventArgs e)
        {
            cboCamChkResultCalib.Text = cboCamChkResultEngCalib.Text;
            if (cboCamChkResultEngCalib.Text == "PASS")
            {
                picCamChkResultEngCalib.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboCamChkResultEngCalib.ForeColor = Color.Green;
            }
            else if (cboCamChkResultEngCalib.Text == "fail")
            {
                picCamChkResultEngCalib.Image = QStations.Properties.Resources.icon_failed_30px;
                cboCamChkResultEngCalib.ForeColor = Color.Red;
            }
            else
            {
                picCamChkResultEngCalib.Image = QStations.Properties.Resources.Input_Off_30px;
                cboCamChkResultEngCalib.ForeColor = Color.Black;
            }
        }

        private void cboProjectChkResultEngCalib_TextChanged(object sender, EventArgs e)
        {
            cboProjectChkResultCalib.Text = cboProjectChkResultEngCalib.Text;
            if (cboProjectChkResultEngCalib.Text == "PASS")
            {
                picProjectChkResultEngCalib.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboProjectChkResultEngCalib.ForeColor = Color.Green;
            }
            else if (cboProjectChkResultEngCalib.Text == "fail")
            {
                picProjectChkResultEngCalib.Image = QStations.Properties.Resources.icon_failed_30px;
                cboProjectChkResultEngCalib.ForeColor = Color.Red;
            }
            else
            {
                picProjectChkResultEngCalib.Image = QStations.Properties.Resources.Input_Off_30px;
                cboProjectChkResultEngCalib.ForeColor = Color.Black;
            }
        }

        private void txtPeakCalibSet_TextChanged(object sender, EventArgs e)
        {
            txtPeakCalib.Text = txtPeakCalibSet.Text;
        }

        private void grpEngCalibration_Leave(object sender, EventArgs e)
        {
            grpEngParameter.Enabled = false; grpEngCalibration.Enabled = false;
            grpEngVCSELSetting.Enabled = false; grpTxEngSetting.Enabled = false; grpEngRxSetting.Enabled = false;
        }

        private void grpEngVCSELSetting_Leave(object sender, EventArgs e)
        {
            grpEngParameter.Enabled = false; grpEngCalibration.Enabled = false;
            grpEngVCSELSetting.Enabled = false; grpTxEngSetting.Enabled = false; grpEngRxSetting.Enabled = false;
        }

        private void grpTxEngSetting_Leave(object sender, EventArgs e)
        {
            grpEngParameter.Enabled = false; grpEngCalibration.Enabled = false;
            grpEngVCSELSetting.Enabled = false; grpTxEngSetting.Enabled = false; grpEngRxSetting.Enabled = false;
        }

        private void grpEngRxSetting_Leave(object sender, EventArgs e)
        {
            grpEngParameter.Enabled = false; grpEngCalibration.Enabled = false;
            grpEngVCSELSetting.Enabled = false; grpTxEngSetting.Enabled = false; grpEngRxSetting.Enabled = false;
        }

        //Engineer-VCSEL Settings
        private void txtLdPwrThrSet_TextChanged(object sender, EventArgs e)
        {
            txtLdPwrThr.Text = txtLdPwrThrSet.Text;
        }

        private void txtLdPwrThrTargetSet_TextChanged(object sender, EventArgs e)
        {
            txtLdPwrThrTarget.Text = txtLdPwrThrTargetSet.Text;
        }

        private void txtLdPwrMaxSet_TextChanged(object sender, EventArgs e)
        {
            txtLdPwrMax.Text = txtLdPwrMaxSet.Text;
        }

        private void txtLdPwrMaxTargetSet_TextChanged(object sender, EventArgs e)
        {
            txtLdPwrMaxTarget.Text = txtLdPwrMaxTargetSet.Text;
        }

        private void txtLdPwrSlopeSet_TextChanged(object sender, EventArgs e)
        {
            txtLdPwrSlope.Text = txtLdPwrSlopeSet.Text;
        }

        private void txtLdPwrSlopeTargetSet_TextChanged(object sender, EventArgs e)
        {
            txtLdPwrSlopeTarget.Text = txtLdPwrSlopeTargetSet.Text;
        }

        private void cboLdPwrThrResultEng_TextChanged(object sender, EventArgs e)
        {
            cboLdPwrThrResult.Text = cboLdPwrThrResultEng.Text;
            if (cboLdPwrThrResultEng.Text == "PASS")
            {
                picLdPwrThrResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrThrResultEng.ForeColor = Color.Green;
                picLdPwrThrResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrThrResult.ForeColor = Color.Green;
            }
            else if (cboLdPwrThrResultEng.Text == "fail")
            {
                picLdPwrThrResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrThrResultEng.ForeColor = Color.Red;
                picLdPwrThrResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrThrResult.ForeColor = Color.Red;
            }
            else
            {
                picLdPwrThrResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrThrResultEng.ForeColor = Color.Black;
                picLdPwrThrResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrThrResult.ForeColor = Color.Black;
            }
        }

        private void cboLdPwrMaxResultEng_TextChanged(object sender, EventArgs e)
        {
            cboLdPwrMaxResult.Text = cboLdPwrMaxResultEng.Text;
            if (cboLdPwrMaxResultEng.Text == "PASS")
            {
                picLdPwrMaxResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrMaxResultEng.ForeColor = Color.Green;
                picLdPwrMaxResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrMaxResult.ForeColor = Color.Green;
            }
            else if (cboLdPwrMaxResultEng.Text == "fail")
            {
                picLdPwrMaxResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrMaxResultEng.ForeColor = Color.Red;
                picLdPwrMaxResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrMaxResult.ForeColor = Color.Red;
            }
            else
            {
                picLdPwrMaxResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrMaxResultEng.ForeColor = Color.Black;
                picLdPwrMaxResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrMaxResult.ForeColor = Color.Black;
            }
        }

        private void cboLdPwrSlopeResultEng_TextChanged(object sender, EventArgs e)
        {
            cboLdPwrSlopeResult.Text = cboLdPwrSlopeResultEng.Text;
            if (cboLdPwrSlopeResultEng.Text == "PASS")
            {
                picLdPwrSlopeResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrSlopeResultEng.ForeColor = Color.Green;
                picLdPwrSlopeResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrSlopeResult.ForeColor = Color.Green;
            }
            else if (cboLdPwrSlopeResultEng.Text == "fail")
            {
                picLdPwrSlopeResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrSlopeResultEng.ForeColor = Color.Red;
                picLdPwrSlopeResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrSlopeResult.ForeColor = Color.Red;
            }
            else
            {
                picLdPwrSlopeResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrSlopeResultEng.ForeColor = Color.Black;
                picLdPwrSlopeResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrSlopeResult.ForeColor = Color.Black;
            }
        }

        private void cboLdPwrStableResultEng_TextChanged(object sender, EventArgs e)
        {
            cboLdPwrStableResult.Text = cboLdPwrStableResultEng.Text;
            if (cboLdPwrStableResultEng.Text == "PASS")
            {
                picLdPwrStableResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrStableResultEng.ForeColor = Color.Green;
                picLdPwrStableResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboLdPwrStableResult.ForeColor = Color.Green;
            }
            else if (cboLdPwrStableResultEng.Text == "fail")
            {
                picLdPwrSlopeResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrStableResultEng.ForeColor = Color.Red;
                picLdPwrStableResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboLdPwrStableResult.ForeColor = Color.Red;
            }
            else
            {
                picLdPwrSlopeResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrStableResultEng.ForeColor = Color.Black;
                picLdPwrStableResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboLdPwrStableResult.ForeColor = Color.Black;
            }
        }

        //Engineer-TxPattern Settings
        private void txtTxPosAngleResultSet_TextChanged(object sender, EventArgs e)
        {
            txtTxPosAngleResult.Text = txtTxPosAngleResultSet.Text;
        }

        private void txtTyPosAngleResultSet_TextChanged(object sender, EventArgs e)
        {
            txtTyPosAngleResult.Text = txtTyPosAngleResultSet.Text;
        }

        private void txtTxDtX13vs100Set_TextChanged(object sender, EventArgs e)
        {
            txtTxDtX13vs100.Text = txtTxDtX13vs100Set.Text;
        }

        private void txtTxDtY13vs100Set_TextChanged(object sender, EventArgs e)
        {
            txtTxDtY13vs100.Text = txtTxDtY13vs100Set.Text;
        }

        private void txtTxRotAngleSet_TextChanged(object sender, EventArgs e)
        {
            txtTxRotAngle.Text = txtTxRotAngleSet.Text;
        }

        private void cboTxPosAngleResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTxPosAngleResult.Text = cboTxPosAngleResultEng.Text;
            if (cboTxPosAngleResultEng.Text == "PASS")
            {
                picTxPosAngleResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxPosAngleResultEng.ForeColor = Color.Green;
                picTxPosAngleResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxPosAngleResult.ForeColor = Color.Green;
            }
            else if (cboTxPosAngleResultEng.Text == "fail")
            {
                picTxPosAngleResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxPosAngleResultEng.ForeColor = Color.Red;
                picTxPosAngleResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxPosAngleResult.ForeColor = Color.Red;
            }
            else
            {
                picTxPosAngleResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxPosAngleResultEng.ForeColor = Color.Black;
                picTxPosAngleResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxPosAngleResult.ForeColor = Color.Black;
            }
        }
        private void cboTyPosAngleResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTyPosAngleResult.Text = cboTyPosAngleResultEng.Text;
            if (cboTyPosAngleResultEng.Text == "PASS")
            {
                picTyPosAngleResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTyPosAngleResultEng.ForeColor = Color.Green;
                picTyPosAngleResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTyPosAngleResult.ForeColor = Color.Green;
            }
            else if (cboTyPosAngleResultEng.Text == "fail")
            {
                picTyPosAngleResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTyPosAngleResultEng.ForeColor = Color.Red;
                picTyPosAngleResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTyPosAngleResult.ForeColor = Color.Red;
            }
            else
            {
                picTyPosAngleResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTyPosAngleResultEng.ForeColor = Color.Black;
                picTyPosAngleResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTyPosAngleResult.ForeColor = Color.Black;
            }
        }
        private void cboTxRotAngleResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTxRotAngleResult.Text = cboTxRotAngleResultEng.Text;
            if (cboTxRotAngleResultEng.Text == "PASS")
            {
                picTxRotAngleResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxRotAngleResultEng.ForeColor = Color.Green;
                picTxRotAngleResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxRotAngleResult.ForeColor = Color.Green;
            }
            else if (cboTxRotAngleResultEng.Text == "fail")
            {
                picTxRotAngleResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxRotAngleResultEng.ForeColor = Color.Red;
                picTxRotAngleResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxRotAngleResult.ForeColor = Color.Red;
            }
            else
            {
                picTxRotAngleResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxRotAngleResultEng.ForeColor = Color.Black;
                picTxRotAngleResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxRotAngleResult.ForeColor = Color.Black;
            }
        }

        private void txtTxFovSet_TextChanged(object sender, EventArgs e)
        {
            txtTxFov.Text = txtTxFovSet.Text;
        }

        private void txtTyFovSet_TextChanged(object sender, EventArgs e)
        {
            txtTyFov.Text = txtTyFovSet.Text;
        }

        private void txtTxSepAngSet_TextChanged(object sender, EventArgs e)
        {
            txtTxSepAng.Text = txtTxSepAngSet.Text;
        }

        private void txtTySepAngSet_TextChanged(object sender, EventArgs e)
        {
            txtTySepAng.Text = txtTySepAngSet.Text;
        }

        private void txtTxAvgSpotSet_TextChanged(object sender, EventArgs e)
        {
            txtTxAvgSpot.Text = txtTxAvgSpotSet.Text;
        }

        private void txtTxAvgFitSet_TextChanged(object sender, EventArgs e)
        {
            txtTxAvgFit.Text = txtTxAvgFitSet.Text;
        }

        private void txtPwrDistributSet_TextChanged(object sender, EventArgs e)
        {
            txtPwrDistribut.Text = txtPwrDistributSet.Text;
        }
        private void cboTxFovResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTxFovResult.Text = cboTxFovResultEng.Text;
            if (cboTxFovResultEng.Text == "PASS")
            {
                picTxFovResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxFovResultEng.ForeColor = Color.Green;
                picTxFovResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxFovResult.ForeColor = Color.Green;
            }
            else if (cboTxFovResultEng.Text == "fail")
            {
                picTxFovResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxFovResultEng.ForeColor = Color.Red;
                picTxFovResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxFovResult.ForeColor = Color.Red;
            }
            else
            {
                picTxFovResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxFovResultEng.ForeColor = Color.Black;
                picTxFovResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxFovResult.ForeColor = Color.Black;
            }
        }

        private void cboTyFovResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTyFovResult.Text = cboTyFovResultEng.Text;
            if (cboTyFovResultEng.Text == "PASS")
            {
                picTyFovResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTyFovResultEng.ForeColor = Color.Green;
                picTyFovResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTyFovResult.ForeColor = Color.Green;
            }
            else if (cboTyFovResultEng.Text == "fail")
            {
                picTyFovResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTyFovResultEng.ForeColor = Color.Red;
                picTyFovResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTyFovResult.ForeColor = Color.Red;
            }
            else
            {
                picTyFovResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTyFovResultEng.ForeColor = Color.Black;
                picTyFovResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTyFovResult.ForeColor = Color.Black;
            }
        }

        private void cboTxSepResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTxSepResult.Text = cboTxSepResultEng.Text;
            if (cboTxSepResultEng.Text == "PASS")
            {
                picTxSepResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxSepResultEng.ForeColor = Color.Green;
                picTxSepResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxSepResult.ForeColor = Color.Green;
            }
            else if (cboTxSepResultEng.Text == "fail")
            {
                picTxSepResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxSepResultEng.ForeColor = Color.Red;
                picTxSepResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxSepResult.ForeColor = Color.Red;
            }
            else
            {
                picTxSepResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxSepResultEng.ForeColor = Color.Black;
                picTxSepResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxSepResult.ForeColor = Color.Black;
            }
        }

        private void cboTySepResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTySepResult.Text = cboTySepResultEng.Text;
            if (cboTySepResultEng.Text == "PASS")
            {
                picTySepResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTySepResultEng.ForeColor = Color.Green;
                picTySepResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTySepResult.ForeColor = Color.Green;
            }
            else if (cboTySepResultEng.Text == "fail")
            {
                picTySepResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTySepResultEng.ForeColor = Color.Red;
                picTySepResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTySepResult.ForeColor = Color.Red;
            }
            else
            {
                picTySepResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTySepResultEng.ForeColor = Color.Black;
                picTySepResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTySepResult.ForeColor = Color.Black;
            }
        }

        private void cboTxAvgSpotResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTxAvgSpotResult.Text = cboTxAvgSpotResultEng.Text;
            if (cboTxAvgSpotResultEng.Text == "PASS")
            {
                picTxAvgSpotResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxAvgSpotResultEng.ForeColor = Color.Green;
                picTxAvgSpotResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxAvgSpotResult.ForeColor = Color.Green;
            }
            else if (cboTxAvgSpotResultEng.Text == "fail")
            {
                picTxAvgSpotResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxAvgSpotResultEng.ForeColor = Color.Red;
                picTxAvgSpotResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxAvgSpotResult.ForeColor = Color.Red;
            }
            else
            {
                picTxAvgSpotResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxAvgSpotResultEng.ForeColor = Color.Black;
                picTxAvgSpotResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxAvgSpotResult.ForeColor = Color.Black;
            }
        }

        private void cboTxAvgFitResultEng_TextChanged(object sender, EventArgs e)
        {
            cboTxAvgFitResult.Text = cboTxAvgFitResultEng.Text;
            if (cboTxAvgFitResultEng.Text == "PASS")
            {
                picTxAvgFitResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxAvgFitResultEng.ForeColor = Color.Green;
                picTxAvgFitResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboTxAvgFitResult.ForeColor = Color.Green;
            }
            else if (cboTxAvgFitResultEng.Text == "fail")
            {
                picTxAvgFitResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxAvgFitResultEng.ForeColor = Color.Red;
                picTxAvgFitResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboTxAvgFitResult.ForeColor = Color.Red;
            }
            else
            {
                picTxAvgFitResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxAvgFitResultEng.ForeColor = Color.Black;
                picTxAvgFitResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboTxAvgFitResult.ForeColor = Color.Black;
            }
        }

        private void cboPwrDistributResultEng_TextChanged(object sender, EventArgs e)
        {
            cboPwrDistributResult.Text = cboPwrDistributResultEng.Text;
            if (cboPwrDistributResultEng.Text == "PASS")
            {
                picPwrDistributResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboPwrDistributResultEng.ForeColor = Color.Green;
                picPwrDistributResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboPwrDistributResult.ForeColor = Color.Green;
            }
            else if (cboPwrDistributResultEng.Text == "fail")
            {
                picPwrDistributResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboPwrDistributResultEng.ForeColor = Color.Red;
                picPwrDistributResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboPwrDistributResult.ForeColor = Color.Red;
            }
            else
            {
                picPwrDistributResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboPwrDistributResultEng.ForeColor = Color.Black;
                picPwrDistributResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboPwrDistributResult.ForeColor = Color.Black;
            }
        }

        //Engineer-RxAssembly Settings
        private void txtRxFovSet_TextChanged(object sender, EventArgs e)
        {
            txtRxFov.Text = txtRxFovSet.Text;
        }

        private void txtRxTxOverlapSet_TextChanged(object sender, EventArgs e)
        {
            txtRxTxOverlap.Text = txtRxTxOverlapSet.Text;
        }

        private void txtRxPatternRotateSet_TextChanged(object sender, EventArgs e)
        {
            txtRxPatternRotate.Text = txtRxPatternRotateSet.Text;
        }

        private void txtRxPathBetwSpotsSet_TextChanged(object sender, EventArgs e)
        {
            txtRxPathBetwSpots.Text = txtRxPathBetwSpotsSet.Text;
        }

        private void cboRxFovResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxFovResult.Text = cboRxFovResultEng.Text;
            if (cboRxFovResultEng.Text == "PASS")
            {
                picRxFovResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxFovResultEng.ForeColor = Color.Green;
                picRxFovResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxFovResult.ForeColor = Color.Green;
            }
            else if (cboRxFovResultEng.Text == "fail")
            {
                picRxFovResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxFovResultEng.ForeColor = Color.Red;
                picRxFovResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxFovResult.ForeColor = Color.Red;
            }
            else
            {
                picRxFovResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxFovResultEng.ForeColor = Color.Black;
                picRxFovResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxFovResult.ForeColor = Color.Black;
            }
        }

        private void cboRxTxOverlapResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxTxOverlapResult.Text = cboRxTxOverlapResultEng.Text;
            if (cboRxTxOverlapResultEng.Text == "PASS")
            {
                picRxTxOverlapResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxTxOverlapResultEng.ForeColor = Color.Green;
                picRxTxOverlapResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxTxOverlapResult.ForeColor = Color.Green;
            }
            else if (cboRxTxOverlapResultEng.Text == "fail")
            {
                picRxTxOverlapResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxTxOverlapResultEng.ForeColor = Color.Red;
                picRxTxOverlapResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxTxOverlapResult.ForeColor = Color.Red;
            }
            else
            {
                picRxTxOverlapResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxTxOverlapResultEng.ForeColor = Color.Black;
                picRxTxOverlapResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxTxOverlapResult.ForeColor = Color.Black;
            }
        }

        private void cboRxPatternRotateResultEng_TextChanged_1(object sender, EventArgs e)
        {
            cboRxPatternRotateResult.Text = cboRxPatternRotateResultEng.Text;
            if (cboRxPatternRotateResultEng.Text == "PASS")
            {
                picRxPatternRotateResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxPatternRotateResultEng.ForeColor = Color.Green;
                picRxPatternRotateResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxPatternRotateResult.ForeColor = Color.Green;
            }
            else if (cboRxPatternRotateResultEng.Text == "fail")
            {
                picRxPatternRotateResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxPatternRotateResultEng.ForeColor = Color.Red;
                picRxPatternRotateResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxPatternRotateResult.ForeColor = Color.Red;
            }
            else
            {
                picRxPatternRotateResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxPatternRotateResultEng.ForeColor = Color.Black;
                picRxPatternRotateResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxPatternRotateResult.ForeColor = Color.Black;
            }
        }

        private void cboRxPathBetwSpotsResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxPathBetwSpotsResult.Text = cboRxPathBetwSpotsResultEng.Text;
            if (cboRxPathBetwSpotsResultEng.Text == "PASS")
            {
                picRxPathBetwSpotsResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxPathBetwSpotsResultEng.ForeColor = Color.Green;
                picRxPathBetwSpotsResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxPathBetwSpotsResult.ForeColor = Color.Green;
            }
            else if (cboRxPathBetwSpotsResultEng.Text == "fail")
            {
                picRxPathBetwSpotsResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxPathBetwSpotsResultEng.ForeColor = Color.Red;
                picRxPathBetwSpotsResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxPathBetwSpotsResult.ForeColor = Color.Red;
            }
            else
            {
                picRxPathBetwSpotsResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxPathBetwSpotsResultEng.ForeColor = Color.Black;
                picRxPathBetwSpotsResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxPathBetwSpotsResult.ForeColor = Color.Black;
            }
        }

        private void txtRxSpotsDetectSet_TextChanged(object sender, EventArgs e)
        {
            txtRxSpotsDetect.Text = txtRxSpotsDetectSet.Text;
        }

        private void txtRxSpotsSizeSet_TextChanged(object sender, EventArgs e)
        {
            txtRxSpotsSize.Text = txtRxSpotsSizeSet.Text;
        }

        private void txtRxSpotPwrDistributSet_TextChanged(object sender, EventArgs e)
        {
            txtRxSpotPwrDistribut.Text = txtRxSpotPwrDistributSet.Text;
        }

        private void txtRxSpotAvgCirSet_TextChanged(object sender, EventArgs e)
        {
            txtRxSpotAvgCir.Text = txtRxSpotAvgCirSet.Text;
        }

        private void cboRxSpotsDetectResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxSpotsDetectResult.Text = cboRxSpotsDetectResultEng.Text;
            if (cboRxSpotsDetectResultEng.Text == "PASS")
            {
                picRxSpotsDetectResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotsDetectResultEng.ForeColor = Color.Green;
                picRxSpotsDetectResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotsDetectResult.ForeColor = Color.Green;
            }
            else if (cboRxSpotsDetectResultEng.Text == "fail")
            {
                picRxSpotsDetectResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotsDetectResultEng.ForeColor = Color.Red;
                picRxSpotsDetectResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotsDetectResult.ForeColor = Color.Red;
            }
            else
            {
                picRxSpotsDetectResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotsDetectResultEng.ForeColor = Color.Black;
                picRxSpotsDetectResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotsDetectResult.ForeColor = Color.Black;
            }
        }

        private void cboRxSpotsSizeResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxSpotsSizeResult.Text = cboRxSpotsSizeResultEng.Text;
            if (cboRxSpotsSizeResultEng.Text == "PASS")
            {
                picRxSpotsSizeResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotsSizeResultEng.ForeColor = Color.Green;
                picRxSpotsSizeResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotsSizeResult.ForeColor = Color.Green;
            }
            else if (cboRxSpotsSizeResultEng.Text == "fail")
            {
                picRxSpotsSizeResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotsSizeResultEng.ForeColor = Color.Red;
                picRxSpotsSizeResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotsSizeResult.ForeColor = Color.Red;
            }
            else
            {
                picRxSpotsSizeResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotsSizeResultEng.ForeColor = Color.Black;
                picRxSpotsSizeResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotsSizeResult.ForeColor = Color.Black;
            }
        }

        private void cboRxSpotPwrDistributResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxSpotPwrDistributResult.Text = cboRxSpotPwrDistributResultEng.Text;
            if (cboRxSpotPwrDistributResultEng.Text == "PASS")
            {
                picRxSpotPwrDistributResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotPwrDistributResultEng.ForeColor = Color.Green;
                picRxSpotPwrDistributResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotPwrDistributResult.ForeColor = Color.Green;
            }
            else if (cboRxSpotPwrDistributResultEng.Text == "fail")
            {
                picRxSpotPwrDistributResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotPwrDistributResultEng.ForeColor = Color.Red;
                picRxSpotPwrDistributResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotPwrDistributResult.ForeColor = Color.Red;
            }
            else
            {
                picRxSpotPwrDistributResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotPwrDistributResultEng.ForeColor = Color.Black;
                picRxSpotPwrDistributResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotPwrDistributResult.ForeColor = Color.Black;
            }
        }

        private void cboRxSpotAvgCirResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxSpotAvgCirResult.Text = cboRxSpotAvgCirResultEng.Text;
            if (cboRxSpotAvgCirResultEng.Text == "PASS")
            {
                picRxSpotAvgCirResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotAvgCirResultEng.ForeColor = Color.Green;
                picRxSpotAvgCirResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSpotAvgCirResult.ForeColor = Color.Green;
            }
            else if (cboRxSpotAvgCirResultEng.Text == "fail")
            {
                picRxSpotAvgCirResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotAvgCirResultEng.ForeColor = Color.Red;
                picRxSpotAvgCirResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSpotAvgCirResult.ForeColor = Color.Red;
            }
            else
            {
                picRxSpotAvgCirResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotAvgCirResultEng.ForeColor = Color.Black;
                picRxSpotAvgCirResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSpotAvgCirResult.ForeColor = Color.Black;
            }
        }

        private void txtRxSensorFpsSet_TextChanged(object sender, EventArgs e)
        {
            txtRxSensorFps.Text = txtRxSensorFpsSet.Text;
        }

        private void txtRxFrameDropsSet_TextChanged(object sender, EventArgs e)
        {
            txtRxFrameDrops.Text = txtRxFrameDropsSet.Text;
        }

        private void txtRxTransSet_TextChanged(object sender, EventArgs e)
        {
            txtRxTrans.Text = txtRxTransSet.Text;
        }

        private void cboRxSensorFpsResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxSensorFpsResult.Text = cboRxSensorFpsResultEng.Text;
            if (cboRxSensorFpsResultEng.Text == "PASS")
            {
                picRxSensorFpsResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSensorFpsResultEng.ForeColor = Color.Green;
                picRxSensorFpsResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxSensorFpsResult.ForeColor = Color.Green;
            }
            else if (cboRxSensorFpsResultEng.Text == "fail")
            {
                picRxSensorFpsResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSensorFpsResultEng.ForeColor = Color.Red;
                picRxSensorFpsResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxSensorFpsResult.ForeColor = Color.Red;
            }
            else
            {
                picRxSensorFpsResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSensorFpsResultEng.ForeColor = Color.Black;
                picRxSensorFpsResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxSensorFpsResult.ForeColor = Color.Black;
            }
        }

        private void cboRxFrameResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxFrameResult.Text = cboRxFrameResultEng.Text;
            if (cboRxFrameResultEng.Text == "PASS")
            {
                picRxFrameResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxFrameResultEng.ForeColor = Color.Green;
                picRxFrameResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxFrameResult.ForeColor = Color.Green;
            }
            else if (cboRxFrameResultEng.Text == "fail")
            {
                picRxFrameResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxFrameResultEng.ForeColor = Color.Red;
                picRxFrameResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxFrameResult.ForeColor = Color.Red;
            }
            else
            {
                picRxFrameResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxFrameResultEng.ForeColor = Color.Black;
                picRxFrameResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxFrameResult.ForeColor = Color.Black;
            }
        }

        private void cboRxTransResultEng_TextChanged(object sender, EventArgs e)
        {
            cboRxTransResult.Text = cboRxTransResultEng.Text;
            if (cboRxTransResultEng.Text == "PASS")
            {
                picRxTransResultEng.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxTransResultEng.ForeColor = Color.Green;
                picRxTransResult.Image = QStations.Properties.Resources.icon_PASS_30px;
                cboRxTransResult.ForeColor = Color.Green;
            }
            else if (cboRxTransResultEng.Text == "fail")
            {
                picRxTransResultEng.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxTransResultEng.ForeColor = Color.Red;
                picRxTransResult.Image = QStations.Properties.Resources.icon_failed_30px;
                cboRxTransResult.ForeColor = Color.Red;
            }
            else
            {
                picRxTransResultEng.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxTransResultEng.ForeColor = Color.Black;
                picRxTransResult.Image = QStations.Properties.Resources.Input_Off_30px;
                cboRxTransResult.ForeColor = Color.Black;
            }
        }

        private void btnPathToFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Please select a folder to save in Auto mode";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    FolderToSaveInAuto = folderDialog.SelectedPath; // 記住選擇的資料夾路徑
                    //MessageBox.Show($"Your selection is : {FolderToSaveInAuto}");
                    txtPathToFolderInAutoMode.Text = FolderToSaveInAuto;
                }
            }
            toolStripStatusLabel1.Text = "Select folder: " + txtPathToFolderInAutoMode.Text;
        }

        //M5.core.OnFrameReceived 的安全呼叫版本 ---參考整合Harris的WinForms範例
        private void OnFrameReceived(byte[] imgBytes)
        {
            // 這一行不要直接 txtLog.AppendText，改用安全版本
            //AppendLog($"OnFrameReceived called, bytes = {imgBytes?.Length}\r\n");

            // 收 frame 這邊是在背景 thread，要用 Invoke 切回 UI
            if (this.IsHandleCreated)
            {
                this.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        using var ms = new MemoryStream(imgBytes);
                        Image frame = Image.FromStream(ms);

                        picM5coreTest.Image?.Dispose(); picM5coreTest.Image = frame;
                        //by Eric:2025-11-27    /*picM5StreamEngParam.Image = null;*/                         
                        picM5StreamEngParam.Image?.Dispose();
                        picM5StreamEngParam.Image = frame;
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"Image decode failed: {ex.Message}\r\n");
                    }
                }));
            }
            else
            {
                // 理論上 Form 已經顯示才會開始收 frame，這裡只是保險
                AppendLog("Handle not created yet when frame received.\r\n");
            }
        }
        private void AppendLog(string text)
        {
            if (this.InvokeRequired)
            {
                // 從背景 thread 叫進來時，轉回 UI thread
                this.BeginInvoke(new Action<string>(AppendLog), text);
                return;
            }

            // 已經在 UI thread，可以安全操作控制項
            txtLog.AppendText(text); txtLogMenu.AppendText(text);
            //toolStripStatusLabel1.Text = text;  //by Eric:2025-11-27 顯示最新狀態在 status bar
        }
        private void UpdateUiByConnectionState(M5ConnectionState state)
        {
            bool connected = state == M5ConnectionState.Connected;

            btnM5StreamOn.Enabled = connected;
            btnM5StreamOff.Enabled = connected;
            btnM5Disconnect.Enabled = connected;
            btnM5LaserOn.Enabled = connected;
            btnM5LaserOff.Enabled = connected;
            btnM5SetLaserPower.Enabled = connected;
            btnM5SetExposure.Enabled = connected;
            btnM5SaveRawFrame.Enabled = connected;
            btnM5SaveImage.Enabled = connected;

            btnM5Connect.Enabled = !connected;

            AppendLog($"[UI] Connection state = {state}\r\n");
        }

        private async void btnConnect_Click(object sender, EventArgs e) //private void...
        {   //btnM5Connect
            AppendLog("Connecting...\r\n"); toolStripStatusLabel1.Text += "Connecting to M5...";

            try
            {
                // 這一段請照你 WPF 的 Btn_Run_M5_Server_Click 把變數搬過來
                string M5_script_dir = @"C:\q_webcam\examples";
                string venv_python = @"C:\q_webcam\venv\Scripts\python.exe";
                string script_path = Path.Combine(M5_script_dir, "stream_webcam_with_manual_control_mgf.py");

                // 這裡看你要不要在 WinForms 也放控件讓使用者輸入 webcam_name / fps / exposure...
                string webcam_name = "V2";  // 先硬寫，之後再拉 TextBox
                int fps = 100;
                int exposure = 1000;    // 200;
                string laser_on = "false"; ; // "true";
                int laser_power = 3;    // 10;

                string args =
                    $"--webcam_name \"{webcam_name}\" " +
                    $"--fps {fps} " +
                    $"--exposure {exposure} " +
                    $"--laser_on {laser_on} " +
                    $"--laser_power {laser_power}";

                // 呼叫 Core 裡真正存在的 API：Start(...)
                _pythonServer.Start(
                    pythonExe: venv_python,
                    scriptPath: script_path,
                    workingDirectory: M5_script_dir,
                    arguments: args,
                    redirectOutput: false  // 如需 log 可改 true
                );

                Thread.Sleep(2000); //delay 3 seconds to wait for server start
                // 再去連 M5 client
                await _m5.ConnectAsync("127.0.0.1", 5001);

                AppendLog("Connected.\r\n"); toolStripStatusLabel1.Text += "...Connected.";

                await M5StreamOff(); //btnStop_Click(null, null); //先停止streaming
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Connect failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                await M5StreamOff();
                AppendLog($"Connect failed: {ex.Message}\r\n"); toolStripStatusLabel1.Text = "Connect failed. You should try to connect again.";
            }
        }

        private async void btnDisconnect_Click(object sender, EventArgs e)
        {
            AppendLog("Disconnecting...\r\n"); toolStripStatusLabel1.Text = "Disconnecting...";

            // 1) 先 try 斷 M5 SDK（等同 WPF 的 Btn_M5_CMD_Quit）
            try
            {
                // WPF 以前是：await _m5Client.DisconnectAsync(sendQuit: true);
                await _m5.DisconnectAsync(sendQuit: true);
                AppendLog("M5 SDK disconnected.\r\n"); toolStripStatusLabel1.Text = "M5 SDK disconnected.";
            }
            catch (Exception ex)
            {
                AppendLog($"M5 disconnect failed: {ex.Message}\r\n"); toolStripStatusLabel1.Text = "M5 disconnect failed.";
            }

            // 2) 再停掉 Python server（等同 WPF 的 Btn_StopPythonServer）
            try
            {
                _pythonServer.Stop();
                AppendLog("Python server stopped.\r\n");
                toolStripStatusLabel1.Text = "Python server stopped.";
            }
            catch (Exception ex)
            {
                AppendLog($"Stop server failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Stop server failed.";
            }
        }

        private async void btnStart_Click(object sender, EventArgs e)   //btnM5StreamOn
        {
            try
            {
                await _m5.SetStreamingAsync(true); chkStreamM5Set.Checked = true;  // 同步勾選狀態
                AppendLog("Start streaming.\r\n"); toolStripStatusLabel1.Text = "Start streaming.";
            }
            catch (Exception ex)
            {
                AppendLog($"Start Streaming failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Start Streaming failed.";
            }
        }

        private async void btnStop_Click(object sender, EventArgs e)    //btnM5StreamOff
        {
            await M5StreamOff();
            /*
            try
            {
                await _m5.SetStreamingAsync(false); chkStreamM5Set.Checked = false;  // 同步勾選狀態
                AppendLog("Stop streaming.\r\n"); //toolStripStatusLabel1.Text = "Stop streaming.";
            }
            catch (Exception ex)
            {
                AppendLog($"Stop Streaming failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Stop Streaming failed.";
            }
            */
        }
        private async Task M5StreamOff()    //btnM5StreamOff
        {
            try
            {
                await _m5.SetStreamingAsync(false); chkStreamM5Set.Checked = false;  // 同步勾選狀態
                AppendLog("Stop streaming.\r\n"); //toolStripStatusLabel1.Text = "Stop streaming.";
            }
            catch (Exception ex)
            {
                AppendLog($"Stop Streaming failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Stop Streaming failed.";
            }
        }

        private async void btnLaserOn_Click(object sender, EventArgs e)
        {
            try
            {
                await _m5.SetLaserAsync(true);  // 等同 WPF 勾選 Checkbox
                AppendLog("Laser ON command sent.\r\n");
                toolStripStatusLabel1.Text = "Laser ON command sent.";
            }
            catch (Exception ex)
            {
                AppendLog($"Laser ON failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Laser ON failed.";
            }
        }
        private async void M5LaserOnAsync()    //btnM5LaserOn
        {
            try
            {
                await _m5.SetLaserAsync(true);  // 等同 WPF 勾選 Checkbox
                AppendLog("Laser ON command sent.\r\n");
                toolStripStatusLabel1.Text = "Laser ON command sent.";
                //return true;
            }
            catch (Exception ex)
            {
                AppendLog($"Laser ON failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Laser ON failed.";
                //return false;
            }
        }

        private /*async*/ void btnLaserOff_Click(object sender, EventArgs e)
        {
            M5LaserOffAsync();
            /*
            try
            {
                await _m5.SetLaserAsync(false);  // 等同 WPF 取消勾選 Checkbox
                AppendLog("Laser OFF command sent.\r\n");
                toolStripStatusLabel1.Text = "Laser OFF command sent.";
            }
            catch (Exception ex)
            {
                AppendLog($"Laser OFF failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Laser OFF failed.";
            }
            */
        }
        private async void M5LaserOffAsync()    //btnM5LaserOff
        {
            try
            {
                await _m5.SetLaserAsync(false);  // 等同 WPF 取消勾選 Checkbox
                AppendLog("Laser OFF command sent.\r\n");
                toolStripStatusLabel1.Text = "Laser OFF command sent.";
                //return true;
            }
            catch (Exception ex)
            {
                AppendLog($"Laser OFF failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "Laser OFF failed.";
                //return false;
            }
        }

        private async void btnSetLaserPower_Click(object sender, EventArgs e)
        {
            if (!int.TryParse(txtLaserPower.Text, out int value))
            {
                AppendLog("請輸入有效的整數雷射功率。\r\n");
                MessageBox.Show("請輸入有效的整數雷射功率。");
                return;
            }
            await LaserPowerSetAsync(value);

            /*
            try
            {
                await _m5.SetLaserPowerAsync(value);
                AppendLog($"Set laser power = {value}\r\n");
                toolStripStatusLabel1.Text = $"Set laser power = {value}";
            }
            catch (Exception ex)
            {
                AppendLog($"Set laser power failed: {ex.Message}\r\n");
                MessageBox.Show($"Set laser power failed: {ex.Message}\r\n");
            }
            */
        }
        private async Task LaserPowerSetAsync(int power)
        {
            try
            {
                await _m5.SetLaserPowerAsync(power);
                AppendLog($"Set laser power = {power}\r\n");
                toolStripStatusLabel1.Text = $"Set laser power = {power}";
            }
            catch (Exception ex)
            {
                AppendLog($"Set laser power failed: {ex.Message}\r\n");
                MessageBox.Show($"Set laser power failed: {ex.Message}\r\n");
            }
        }

        private async void btnSetExposure_Click(object sender, EventArgs e)
        {
            if (!int.TryParse(txtExposure.Text, out int value))
            {
                AppendLog("請輸入有效的曝光值。\r\n"); MessageBox.Show("請輸入有效的曝光值。");
                return;
            }

            try
            {
                await _m5.SetExposureAsync(value);
                AppendLog($"Set exposure = {value}\r\n"); toolStripStatusLabel1.Text = $"Set exposure = {value}";
            }
            catch (InvalidOperationException ex)
            {
                // 這裡通常是「未連線」
                AppendLog($"SetExposure blocked: {ex.Message}\r\n");
                MessageBox.Show($"SetExposure blocked: {ex.Message}\r\n");
            }
            catch (Exception ex)
            {
                AppendLog($"SetExposure failed: {ex.Message}\r\n");
                MessageBox.Show($"SetExposure failed: {ex.Message}\r\n");
            }
        }

        private async void btnSaveRawFrame_Click(object sender, EventArgs e)
        {
            try
            {
                await _m5.SaveFrameAsync(); //save in 'c:\q_webcam\examples' folder as .raw file
                AppendLog("SaveFrame command sent.\r\n");
                toolStripStatusLabel1.Text = "SaveFrame command sent.";
            }
            catch (Exception ex)
            {
                AppendLog($"SaveFrame failed: {ex.Message}\r\n");
                toolStripStatusLabel1.Text = "SaveFrame failed.";
            }
        }

        //---M5 AlgorithmEngine PD LIV & Stability Test---//
        private void LoadImageToPictureBox(PictureBox pictureBox, string sourcePath, string tempFileName)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                return;

            try
            {
                string tempPng = Path.Combine(Path.GetTempPath(), tempFileName);
                File.Copy(sourcePath, tempPng, overwrite: true);

                // 避免檔案被鎖定，先釋放舊圖，再用 FileStream 讀
                if (pictureBox.Image != null)
                {
                    pictureBox.Image.Dispose();
                    pictureBox.Image = null;
                }

                using (var fs = new FileStream(tempPng, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    pictureBox.Image = Image.FromStream(fs);
                }
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrEmpty(txtError.Text))
                    txtError.AppendText(Environment.NewLine);
                txtError.AppendText($"[LoadImage Error] {ex.Message}");
            }
        }

        private async void btnRunAll_Click(object sender, EventArgs e)
        {
            // 避免重複觸發，可以暫時 disable 按鈕
            btnRunAll.Enabled = false;
            btnRunLivTest.Enabled = false;
            btnRunStabilityTest.Enabled = false;

            picLivPlot.Image = null; picStabilityPlot.Image = null;
            try
            {
                var testCurrs = new List<double> { 0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0 };
                var testPwrss = new List<double> { 0.0, 0.0, 1.1, 1.6, 2.1, 2.6, 3.1, 3.6, 4.1, 4.6, 5.1 };

                await RunLivTestAsync(testCurrs, testPwrss);    //await RunLivTestAsync();
                var samplePowerArray = new List<double>
                {
                12.80, 12.81, 12.79, 12.80, 12.82, 12.78, 12.80, 12.81, 12.79, 12.80,
                12.81, 12.82, 12.79, 12.78, 12.80, 12.81, 12.83, 12.79, 12.80, 12.81,
                12.78, 12.80, 12.82, 12.81, 12.79, 12.80, 12.81, 12.79, 12.80, 12.82
                };
                await RunStabilityTestAsync(samplePowerArray);  //await RunStabilityTestAsync();
            }
            finally
            {
                btnRunAll.Enabled = true;
                btnRunLivTest.Enabled = true;
                btnRunStabilityTest.Enabled = true;
            }
        }

        private /*async*/ void btnM5RunLiv_Click(object sender, EventArgs e)
        {
            btnLivChartClear_Click(sender, e); //clear chart first
            btnRunLiv.Enabled = false; btnRunStability.Enabled = false;
            btnLivChartClear.Enabled = false; btnStableChartClear.Enabled = false;
            toolStripStatusLabel1.Text = "M5 LIV Test Running...";
            lblLivStatus.Text = "Running..."; txtCalculatedLiv.Text = "Running...";

            M5LaserOnAsync();

            List<double> dblM5CurrentList = new List<double>();
            List<double> dblM5PwrList = new List<double>();
            for (int value = 0; value <= 10; value += 1)
            {
                Task taskLivLaserPwrSet = LaserPowerSetAsync(value);
                Thread.Sleep(500);
                dblM5CurrentList.Add((double)value);
                double M5Pwr = Math.Round(M5PowerRead(value) * 1000, 3);    //to mW
                Thread.Sleep(500); //wait for power meter read
                dblM5PwrList.Add(M5Pwr);
            }
            // (可選) 顯示 list 中的內容 (用於驗證)
            txtLivValList.Text = "itemCurr : \r\n";
            foreach (double itemCurr in dblM5CurrentList) { Debug.WriteLine(itemCurr); txtLivValList.Text += itemCurr.ToString() + "\r\n"; }
            txtLivValList.Text += "\r\nitemPwr : \r\n";
            foreach (double itemPwr in dblM5PwrList) { Console.WriteLine(itemPwr); txtLivValList.Text += itemPwr.ToString() + "\r\n"; }
            //Task taskLIV = RunLivTestAsync(dblM5CurrentList, dblM5PwrList);
            try
            {
                Task taskLIV = RunLivTestAsync(dblM5CurrentList, dblM5PwrList);
                //taskLIV.Wait();
            }
            finally
            {
                btnRunLiv.Enabled = true; btnRunStability.Enabled = true;
                btnLivChartClear.Enabled = true; btnStableChartClear.Enabled = true;
                toolStripStatusLabel1.Text = "M5 LIV Test Completed.";
                M5LaserOffAsync();
            }
            
        }

        private async void btnRunStability_Click(object sender, EventArgs e)
        {
            btnStableChartClear_Click(null, null);
            btnRunAll.Enabled = false;
            btnRunLivEng.Enabled = false; btnRunLiv.Enabled = false;
            btnRunStabilityEng.Enabled = false; btnRunStability.Enabled = false;
            btnLivChartClear.Enabled = false; btnStableChartClear.Enabled = false;
            lblStabilityStatus.Text = "Running...";
            txtCalculatedStability.Text = "Running...";

            M5LaserOnAsync();

            //int iCurrentVal = 7; //設置穩定性測試的雷射功率值
            if (txtLaserPower.Text == "") { txtLaserPower.Text = "7"; }

            if (!int.TryParse(txtLaserPower.Text, out int iCurrentVal))
            {
                AppendLog("請輸入有效的整數雷射功率。\r\n");
                MessageBox.Show("請輸入有效的整數雷射功率。");
                btnRunLivEng.Enabled = true; btnRunStabilityEng.Enabled = true;
                btnLivChartClearEng.Enabled = true; btnStableChartClearEng.Enabled = true;
                toolStripStatusLabel1.Text = "請輸入有效的整數雷射功率。";
                return;
            }
            var dblM5StablePwrList = new List<double>();
            await LaserPowerSetAsync(iCurrentVal);
            Thread.Sleep(500);
            for (int i = 0; i < 30; i += 1)
            {
                double M5Pwr = Math.Round(MeasurePower() * 1000, 3);   //M5PowerRead(iCurrentVal);
                Thread.Sleep(300); //wait for power meter read
                dblM5StablePwrList.Add(M5Pwr);
                //Thread.Sleep(200);
            }
            //foreach (double item in dblM5StablePwrList) { Console.WriteLine(item); }
            
            try
            {
                await RunStabilityTestAsync(dblM5StablePwrList);  //await RunStabilityTestAsync();
            }
            finally
            {
                btnRunAll.Enabled = true;
                btnRunLivEng.Enabled = true; btnRunLiv.Enabled = true;
                btnRunStabilityEng.Enabled = true; btnRunStability.Enabled = true;
                btnLivChartClear.Enabled = true; btnStableChartClear.Enabled = true;
                M5LaserOffAsync();
            }

        }
        private async void btnRunStabilityTest_Click(object sender, EventArgs e)
        {
            btnRunAll.Enabled = false;
            btnRunLivTest.Enabled = false;
            btnRunStabilityTest.Enabled = false;
            var samplePowerArray = new List<double> //模擬的測試資料
            {
                12.80, 12.81, 12.79, 12.80, 12.82, 12.78, 12.80, 12.81, 12.79, 12.80,
                12.81, 12.82, 12.79, 12.78, 12.80, 12.81, 12.83, 12.79, 12.80, 12.81,
                12.78, 12.80, 12.82, 12.81, 12.79, 12.80, 12.81, 12.79, 12.80, 12.82
            };
            try
            {
                await RunStabilityTestAsync(samplePowerArray);  //await RunStabilityTestAsync();
            }
            finally
            {
                btnRunAll.Enabled = true;
                btnRunLivTest.Enabled = true;
                btnRunStabilityTest.Enabled = true;
            }
        }

        // 1) LIV 測試核心邏輯
        private async Task RunLivTestAsync(List<double> currents, List<double> powers)
        {
            lblLivStatusEng.Text = "Running..."; lblLivStatusTest.Text = "Running..."; lblLivStatus.Text = "Running...";
            // 這裡放你現在 btnRunLiv_Click 裡的內容，
            // 把 sender/e 相關的東西拿掉，改成純邏輯
            // 測試資料，可以先複製 WPF 的
            //var currents = new List<double> { 0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0 };
            //var powers = new List<double> { 0.0, 0.0, 1.1, 1.6, 2.1, 2.6, 3.1, 3.6, 4.1, 4.6, 5.1 };

            double thresholdPower = double.Parse(txtLdPwrThrTargetSet.Text);//1.0;

            var rules = new PdLivThresholdRules
            {
                SlopeLSL = 0.6,
                MaxPowerLSL = 11.0,
                ThresholdPowerLSL = 1.0,
                ThresholdPowerUSL = 1.5
            };

            try
            {
                var result = await _algoEngine.RunPdLivAsync(
                    currents,
                    powers,
                    thresholdPower,
                    rules);

                result.ThrowIfError();

                lblLivStatusEng.Text = result.PassStatus ? "PASS" : "FAIL";
                lblLivStatusTest.Text = result.PassStatus ? "PASS" : "FAIL";
                lblLivStatus.Text = result.PassStatus ? "PASS" : "FAIL";

                if (result.CalculatedValues != null)
                {
                    var sb = new StringBuilder();
                    foreach (var kv in result.CalculatedValues)
                    {
                        sb.AppendLine($"{kv.Key} = {kv.Value}");
                    }
                    txtCalculatedLivTest.Text = sb.ToString();
                    txtCalculatedLivEng.Text = sb.ToString(); txtCalculatedLiv.Text = sb.ToString();
                }

                LoadImageToPictureBox(picLivPlot, result.PlotPath, "pd_plot_liv_display.png");
                picLivChartEng.SizeMode = PictureBoxSizeMode.Zoom;
                LoadImageToPictureBox(picLivChartEng, result.PlotPath, "pd_plot_liv_display.png");
                picLivChart.SizeMode = PictureBoxSizeMode.Zoom;
                LoadImageToPictureBox(picLivChart, result.PlotPath, "pd_plot_liv_display.png");
            }
            catch (InvalidOperationException ex)
            {
                lblLivStatusEng.Text = "ERROR"; lblLivStatusTest.Text = "ERROR"; lblLivStatus.Text = "ERROR";
                if (!string.IsNullOrEmpty(txtError.Text))
                    txtError.AppendText(Environment.NewLine);
                txtError.AppendText($"[PD LIV Error] {ex.Message}");
            }
            catch (Exception ex)
            {
                lblLivStatusEng.Text = "ERROR"; lblLivStatusTest.Text = "ERROR"; lblLivStatus.Text = "ERROR";
                if (!string.IsNullOrEmpty(txtError.Text))
                    txtError.AppendText(Environment.NewLine);
                txtError.AppendText($"[PD LIV Exception] {ex}");
            }
            btnRunLivEng.Enabled = true; btnRunStabilityEng.Enabled = true;
            btnLivChartClearEng.Enabled = true; btnStableChartClearEng.Enabled = true;
            toolStripStatusLabel1.Text = "PD LIV Test Completed.";
        }
        private void btnLivChartClear_Click(object sender, EventArgs e)
        {
            picLivChartEng.Image = null; picLivChart.Image = null;
            lblLivStatusEng.Text = ""; txtCalculatedLivEng.Text = ""; txtCalculatedLiv.Text = ""; lblLivStatus.Text = "";
        }

        // 2) Stability 測試核心邏輯
        private async Task RunStabilityTestAsync(List<double> samplePowerArray)  //RunStabilityTestAsync()
        {
            // 這裡放你現在 btnRunStability_Click 裡的內容
            //var samplePowerArray = new List<double>
            //{
            //    12.80, 12.81, 12.79, 12.80, 12.82, 12.78, 12.80, 12.81, 12.79, 12.80,
            //    12.81, 12.82, 12.79, 12.78, 12.80, 12.81, 12.83, 12.79, 12.80, 12.81,
            //    12.78, 12.80, 12.82, 12.81, 12.79, 12.80, 12.81, 12.79, 12.80, 12.82
            //};

            double stabilityUSL = 1.0;

            lblStabilityStatusEng.Text = "Running..."; lblStabilityStatus.Text = "Running...";
            lblStabilityStatusTest.Text = "Running...";

            try
            {
                var result = await _algoEngine.RunPdStabilityAsync(
                    samplePowerArray,
                    stabilityUSL);

                result.ThrowIfError();

                lblStabilityStatusEng.Text = result.PassStatus ? "PASS" : "FAIL";
                lblStabilityStatus.Text = result.PassStatus ? "PASS" : "FAIL";
                lblStabilityStatusTest.Text = result.PassStatus ? "PASS" : "FAIL";

                if (result.CalculatedValues != null)
                {
                    var sb = new StringBuilder();
                    foreach (var kv in result.CalculatedValues)
                    {
                        sb.AppendLine($"{kv.Key} = {kv.Value}");
                    }
                    txtCalculatedStabilityEng.Text = sb.ToString();
                    txtCalculatedStabilityTest.Text = sb.ToString();
                    txtCalculatedStability.Text = sb.ToString();
                }

                LoadImageToPictureBox(picStabilityPlot, result.PlotPath, "pd_plot_stability_display.png");

                picPwrTimeChartEng.SizeMode = PictureBoxSizeMode.Zoom;
                LoadImageToPictureBox(picPwrTimeChartEng, result.PlotPath, "pd_plot_stability_display.png");
                picPwrTimeChart.SizeMode = PictureBoxSizeMode.Zoom;
                LoadImageToPictureBox(picPwrTimeChart, result.PlotPath, "pd_plot_stability_display.png");
            }
            catch (InvalidOperationException ex)
            {
                lblStabilityStatusEng.Text = "ERROR"; lblStabilityStatusTest.Text = "ERROR"; lblStabilityStatus.Text = "ERROR";
                if (!string.IsNullOrEmpty(txtError.Text))
                    txtError.AppendText(Environment.NewLine);
                txtError.AppendText($"[PD Stability Error] {ex.Message}");
            }
            catch (Exception ex)
            {
                lblStabilityStatusEng.Text = "ERROR"; lblStabilityStatusTest.Text = "ERROR"; lblStabilityStatus.Text = "ERROR";
                if (!string.IsNullOrEmpty(txtError.Text))
                    txtError.AppendText(Environment.NewLine);
                txtError.AppendText($"[PD Stability Exception] {ex}");
            }
        }

        private void btnStableChartClear_Click(object sender, EventArgs e)
        {
            picPwrTimeChartEng.Image = null; picPwrTimeChart.Image = null;
            lblStabilityStatusEng.Text = string.Empty; txtCalculatedStabilityEng.Text = string.Empty;
            lblStabilityStatus.Text = string.Empty; txtCalculatedStability.Text = string.Empty;
        }

        private async void btnM5LivTest_Click(object sender, EventArgs e) //LIV test
        {
            btnRunAll.Enabled = false;
            btnRunLivTest.Enabled = false;
            btnRunStabilityTest.Enabled = false;
            try
            {
                var testCurrs = new List<double> { 0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0 };
                var testPwrss = new List<double> { 0.0, 0.0, 1.1, 1.6, 2.1, 2.6, 3.1, 3.6, 4.1, 4.6, 5.1 };
                await RunLivTestAsync(testCurrs, testPwrss);    //await RunLivTestAsync();
            }
            finally
            {
                btnRunAll.Enabled = true;
                btnRunLivTest.Enabled = true;
                btnRunStabilityTest.Enabled = true;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you sure to exit？", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                btnLaserOff_Click(sender, e); //turn off laser when exit
                _m5.Dispose();
                Thread.Sleep(3000); //wait for laser off command send complete
            }
            else { e.Cancel = true; }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            btnDisconnect_Click(sender, e);
            _m5.QuitAsync();
            //Environment.Exit(Environment.ExitCode);
        }

        private void btnPwrReadTest_Click(object sender, EventArgs e)
        {
            if (!int.TryParse(txtLaserPower.Text, out int value))
            {
                AppendLog("請輸入有效的整數雷射功率。\r\n");
                MessageBox.Show("請輸入有效的整數雷射功率。");
                return;
            }
            Task PwrReadtask = LaserPowerSetAsync(value);
            /*
            try
            {
                _m5.SetLaserPowerAsync(value);
                AppendLog($"Set laser power = {value}\r\n");
                toolStripStatusLabel1.Text = $"Set laser power = {value}";
            }
            catch (Exception ex)
            {
                AppendLog($"Set laser power failed: {ex.Message}\r\n");
                MessageBox.Show($"Set laser power failed: {ex.Message}\r\n");
            }
            */

            //M5PowerRead(value);
            labelPower.Text = (M5PowerRead(value) * 1000).ToString() + " mW";   // Display power in mW
        }
        private double M5PowerRead(double power)
        {
            try
            {
                double powerValue = MeasurePower();
                //int err = pm100d.measPower(out powerValue);                
                //labelPower.Text = (powerValue * 1000).ToString() + " mW";   // Display power in mW
                return powerValue;   //power in uW
            }
            catch (BadImageFormatException bie)
            {
                AppendLog(bie.Message);
            }
            catch (NullReferenceException nre)
            {
                AppendLog(nre.Message);
            }
            catch (ExternalException ex)
            {
                AppendLog(ex.Message);
                MessageBox.Show("Connect to Power Meter failed.\r\n" + ex.Message);
            }
            finally
            {
                if (pm100d != null)
                    pm100d.Dispose();
            }
            //toolStripStatusLabel1.Text = "Power Meter Read Error.";
            return -1;
        }

        private async void btnRunLivTest_Click(object sender, EventArgs e)
        {
            btnRunAll.Enabled = false;
            btnRunLivEng.Enabled = false;
            btnRunStabilityEng.Enabled = false;
            try
            {
                var testCurrs = new List<double> { 0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0 };
                var testPwrss = new List<double> { 0.0, 0.0, 1.1, 1.6, 2.1, 2.6, 3.1, 3.6, 4.1, 4.6, 5.1 };
                await RunLivTestAsync(testCurrs, testPwrss);    //await RunLivTestAsync();
            }
            finally
            {
                btnRunAll.Enabled = true;
                btnRunLivEng.Enabled = true;
                btnRunStabilityEng.Enabled = true;
            }
        }

        private void BtnInitialCard_Click(object sender, EventArgs e)
        {
            ushort uCount = 0, uCardNo = 0;
            g_uRet = CEtherCAT_DLL.CS_ECAT_Master_Open(ref g_nESCExistCards);
            g_bInitialFlag = false;
            CmbCardNo.Items.Clear();
            if (g_nESCExistCards == 0)
            {
                AddErrMsg("No EtherCat can be found!", true);
            }
            else
            {
                for (uCount = 0; uCount < 32; uCount++)
                {
                    g_uESCCardNoList[uCount] = 99;
                }

                for (uCount = 0; uCount < g_nESCExistCards; uCount++)
                {
                    g_uRet = CEtherCAT_DLL.CS_ECAT_Master_Get_CardSeq(uCount, ref uCardNo);
                    g_uRet = CEtherCAT_DLL.CS_ECAT_Master_Initial(uCardNo);
                    if (g_uRet != 0)
                    {
                        AddErrMsg("_ECAT_Master_Initial, ErrorCode = " + g_uRet.ToString(), true);
                    }
                    else
                    {
                        g_uESCCardNoList[uCount] = uCardNo;
                        CmbCardNo.Items.Add(uCardNo.ToString());
                        g_bInitialFlag = true;
                    }
                }

                if (g_bInitialFlag == true)
                {
                    CmbCardNo.SelectedIndex = 0;
                    g_uESCCardNo = g_uESCCardNoList[0];
                }
            }
        }
        private void AddErrMsg(string strMsg, bool bShowErr = false)
        {
            LstErrMsg.SelectedIndex = LstErrMsg.Items.Add(strMsg);

            if (bShowErr == true)
                MessageBox.Show(strMsg);
        }
        private void BtnFindSlave_Click(object sender, EventArgs e)
        {
            short nSID = 0, Cnt = 0;
            ushort uNID = 0, uSlaveNum = 0, uReMapNodeID = 0;
            uint uVendorID = 0, uProductCode = 0, uRevisionNo = 0, uSlaveDCTime = 0;
            string strMsg;
            TxtSlaveNum.Text = "0";
            CmbNode1.Items.Clear();
            CmbNodeID1.Items.Clear();
            CmbSlotID1.Items.Clear();
            g_uRet = CEtherCAT_DLL.CS_ECAT_Master_Get_SlaveNum(g_uESCCardNo, ref uSlaveNum);

            if (g_uRet != CEtherCAT_DLL_Err.ERR_ECAT_NO_ERROR)
            {
                AddErrMsg("_ECAT_Master_Get_SlaveNum, ErrorCode = " + g_uRet.ToString(), true);
            }
            else
            {
                if (uSlaveNum == 0)
                {
                    AddErrMsg("Card NO: " + g_uESCCardNo.ToString() + " No slave found!", true);
                }
                else
                {
                    CmbNode1.Items.Clear();
                    CmbNodeID1.Items.Clear();
                    CmbSlotID1.Items.Clear();
                    CmbNode2.Items.Clear();
                    CmbNodeID2.Items.Clear();
                    CmbSlotID2.Items.Clear();
                    CmbNode3.Items.Clear();
                    CmbNodeID3.Items.Clear();
                    CmbSlotID3.Items.Clear();
                    TxtSlaveNum.Text = uSlaveNum.ToString();
                    for (uNID = 0; uNID < uSlaveNum; uNID++)
                    {
                        g_uRet = CEtherCAT_DLL.CS_ECAT_Master_Get_Slave_Info(g_uESCCardNo, uNID, ref uReMapNodeID, ref uVendorID, ref uProductCode, ref uRevisionNo, ref uSlaveDCTime);

                        if (uVendorID == 0x1DD && uProductCode == 0x10305070) // A2E
                        {
                            nSID = 0;
                            strMsg = "NodeID:" + uNID + " - SlotID:" + nSID + "-A2E";
                            CmbNode1.Items.Add(strMsg);
                            CmbNodeID1.Items.Add(uNID.ToString());
                            CmbSlotID1.Items.Add(nSID.ToString());

                            CmbNode2.Items.Add(strMsg);
                            CmbNodeID2.Items.Add(uNID.ToString());
                            CmbSlotID2.Items.Add(nSID.ToString());

                            CmbNode3.Items.Add(strMsg);
                            CmbNodeID3.Items.Add(uNID.ToString());
                            CmbSlotID3.Items.Add(nSID.ToString());
                            Cnt++;
                        }
                        else if ((uVendorID == 0x1A05 || uVendorID == 0x1DD) && uProductCode == 0x0624) //Ec4Axis
                        {
                            for (nSID = 0; nSID < 4; nSID++)
                            {
                                strMsg = "NodeID:" + uNID + " - SlotID:" + nSID + "-Ec4Axis";
                                CmbNode1.Items.Add(strMsg);
                                CmbNodeID1.Items.Add(uNID.ToString());
                                CmbSlotID1.Items.Add(nSID.ToString());

                                CmbNode2.Items.Add(strMsg);
                                CmbNodeID2.Items.Add(uNID.ToString());
                                CmbSlotID2.Items.Add(nSID.ToString());

                                CmbNode3.Items.Add(strMsg);
                                CmbNodeID3.Items.Add(uNID.ToString());
                                CmbSlotID3.Items.Add(nSID.ToString());
                                Cnt++;
                            }
                        }
                        else if ((uVendorID == 0x1A05 || uVendorID == 0x1DD) && uProductCode == 0x5621) //EcAxis
                        {
                            nSID = 0;
                            strMsg = "NodeID:" + uNID + " - SlotID:" + nSID + "-EcAxis";
                            CmbNode1.Items.Add(strMsg);
                            CmbNodeID1.Items.Add(uNID.ToString());
                            CmbSlotID1.Items.Add(nSID.ToString());

                            CmbNode2.Items.Add(strMsg);
                            CmbNodeID2.Items.Add(uNID.ToString());
                            CmbSlotID2.Items.Add(nSID.ToString());

                            CmbNode3.Items.Add(strMsg);
                            CmbNodeID3.Items.Add(uNID.ToString());
                            CmbSlotID3.Items.Add(nSID.ToString());
                            Cnt++;
                        }
                        else if (uVendorID == 0x539 && uProductCode == 0x2200001) //Yaskawa
                        {
                            nSID = 0;
                            strMsg = "NodeID:" + uNID + " - SlotID:" + nSID + "-Yaskawa";
                            CmbNode1.Items.Add(strMsg);
                            CmbNodeID1.Items.Add(uNID.ToString());
                            CmbSlotID1.Items.Add(nSID.ToString());

                            CmbNode2.Items.Add(strMsg);
                            CmbNodeID2.Items.Add(uNID.ToString());
                            CmbSlotID2.Items.Add(nSID.ToString());

                            CmbNode3.Items.Add(strMsg);
                            CmbNodeID3.Items.Add(uNID.ToString());
                            CmbSlotID3.Items.Add(nSID.ToString());
                            Cnt++;
                        }
                    }

                    if (Cnt > 0)
                    {
                        CmbNode1.SelectedIndex = 0;
                        CmbNodeID1.SelectedIndex = 0;
                        CmbSlotID1.SelectedIndex = 0;

                        BtnMotorStop.Enabled = true;
                        BtnMoveLeft.Enabled = true;
                        BtnMoveRight.Enabled = true;
                        RdoSVOFF.Enabled = true;
                        RdoSVON.Enabled = true;
                        BtnResetAlarm.Enabled = true;
                        BtnResetStatus.Enabled = true;
                        BtnChangeDist.Enabled = true;
                        BtnChangeVel.Enabled = true;
                        ChkSetLimit.Enabled = true;
                        ChkSetGear.Enabled = true;
                        TrcFeedrate.Enabled = true;
                    }
                }
            }
        }
        private void RdoSVON_CheckedChanged(object sender, EventArgs e)
        {
            ushort uCheckOnOff = 0;
            string strMsg;
            if (RdoSVON.Checked == true)
                uCheckOnOff = 1;
            g_uRet = CEtherCAT_DLL.CS_ECAT_Slave_Motion_Set_Svon(g_uESCCardNo, g_uESCNodeID[0], g_uESCSlotID[0], uCheckOnOff);

            if (g_uRet != CEtherCAT_DLL_Err.ERR_ECAT_NO_ERROR)
            {
                strMsg = "CS_ECAT_Slave_Motion_Set_Svon, Error Code = " + g_uRet.ToString();
                AddErrMsg(strMsg, true);
            }

            if (g_nSelectMode > 2)
            {
                g_uRet = CEtherCAT_DLL.CS_ECAT_Slave_Motion_Set_Svon(g_uESCCardNo, g_uESCNodeID[1], g_uESCSlotID[1], uCheckOnOff);

                if (g_uRet != CEtherCAT_DLL_Err.ERR_ECAT_NO_ERROR)
                {
                    strMsg = "CS_ECAT_Slave_Motion_Set_Svon, Error Code = " + g_uRet.ToString();
                    AddErrMsg(strMsg, true);
                }
            }

            if (g_nSelectMode > 8)
            {
                g_uRet = CEtherCAT_DLL.CS_ECAT_Slave_Motion_Set_Svon(g_uESCCardNo, g_uESCNodeID[2], g_uESCSlotID[2], uCheckOnOff);

                if (g_uRet != CEtherCAT_DLL_Err.ERR_ECAT_NO_ERROR)
                {
                    strMsg = "CS_ECAT_Slave_Motion_Set_Svon, Error Code = " + g_uRet.ToString();
                    AddErrMsg(strMsg, true);
                }
            }
        }
        private void RdoCSPMode01_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 1;
            SetParamInfo(1, "Distance", "pluse");
            SetParamInfo(2, "StartVel", "pps");
            SetParamInfo(3, "MaxVel", "pps");
            SetParamInfo(4, "EndVel", "pps");
            SetParamInfo(5, "Acc.", "sec", "0.1");
            SetParamInfo(6, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, false);
            EnableAxis(3, false);
        }

        private void RdoCSPMode02_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 2;
            SetParamInfo(1, "StartVel", "pps");
            SetParamInfo(2, "MaxVel", "pps");
            SetParamInfo(3, "Acc.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, false);
            EnableAxis(3, false);
        }

        private void RdoCSPMode03_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 3;
            SetParamInfo(1, "Distance X", "pluse");
            SetParamInfo(2, "Distance Y", "pluse");
            SetParamInfo(3, "StartVel", "pps");
            SetParamInfo(4, "MaxVel", "pps");
            SetParamInfo(5, "EndVel", "pps");
            SetParamInfo(6, "Acc.", "sec", "0.1");
            SetParamInfo(7, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, true);
            EnableAxis(3, false);
        }

        private void RdoCSPMode04_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 4;
            SetParamInfo(1, "Center X", "pluse");
            SetParamInfo(2, "Center Y", "pluse");
            SetParamInfo(3, "Angle", "360/cir");
            SetParamInfo(4, "StartVel", "pps");
            SetParamInfo(5, "MaxVel", "pps");
            SetParamInfo(6, "EndVel", "pps");
            SetParamInfo(7, "Acc.", "sec", "0.1");
            SetParamInfo(8, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, true);
            EnableAxis(3, false);
        }

        private void RdoCSPMode05_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 5;
            SetParamInfo(1, "End X", "pluse");
            SetParamInfo(2, "End Y", "pluse");
            SetParamInfo(3, "Angle", "360/cir");
            SetParamInfo(4, "StartVel", "pps");
            SetParamInfo(5, "MaxVel", "pps");
            SetParamInfo(6, "EndVel", "pps");
            SetParamInfo(7, "Acc.", "sec", "0.1");
            SetParamInfo(8, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, true);
            EnableAxis(3, false);
        }

        private void RdoCSPMode06_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 6;
            SetParamInfo(1, "Center X", "pluse");
            SetParamInfo(2, "Center Y", "pluse");
            SetParamInfo(3, "End X", "pluse");
            SetParamInfo(4, "End Y", "pluse");
            SetParamInfo(5, "StartVel", "pps");
            SetParamInfo(6, "MaxVel", "pps");
            SetParamInfo(7, "EndVel", "pps");
            SetParamInfo(8, "Acc.", "sec", "0.1");
            SetParamInfo(9, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, true);
            EnableAxis(3, false);
        }

        private void RdoCSPMode07_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 7;
            SetParamInfo(1, "Center X", "pluse");
            SetParamInfo(2, "Center Y", "pluse");
            SetParamInfo(3, "Interval", "pluse");
            SetParamInfo(4, "Angle", "pluse");
            SetParamInfo(5, "StartVel", "pps");
            SetParamInfo(6, "MaxVel", "pps");
            SetParamInfo(7, "EndVel", "pps");
            SetParamInfo(8, "Acc.", "sec", "0.1");
            SetParamInfo(9, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, true);
            EnableAxis(3, false);
        }

        private void RdoCSPMode08_CheckedChanged(object sender, EventArgs e)
        {
            HideAllParam();
            g_nSelectMode = 8;
            SetParamInfo(1, "Center X", "pluse");
            SetParamInfo(2, "Center Y", "pluse");
            SetParamInfo(3, "End X", "pluse");
            SetParamInfo(4, "End Y", "pluse");
            SetParamInfo(5, "CycleNum", "cir.");
            SetParamInfo(6, "StartVel", "pps");
            SetParamInfo(7, "MaxVel", "pps");
            SetParamInfo(8, "EndVel", "pps");
            SetParamInfo(9, "Acc.", "sec", "0.1");
            SetParamInfo(10, "Dec.", "sec", "0.1");

            EnableAxis(1, true);
            EnableAxis(2, true);
            EnableAxis(3, false);
        }
        private void HideAllParam()
        { // 隱藏所有參數欄 再一一開啟
            for (int nNo = 1; nNo < 12; nNo++)
            {
                g_pTxtParam[nNo].Visible = false;
                g_pLabParamTitle[nNo].Visible = false;
                g_pLabParamUnit[nNo].Visible = false;
            }
        }
        private void SetParamInfo(int nParamNo, string strParamTitle, string strParamUnit, string strValue = "0")
        {
            if (nParamNo < 1 || nParamNo > 11) return;
            g_pLabParamTitle[nParamNo].Visible = true;
            g_pLabParamUnit[nParamNo].Visible = true;
            g_pTxtParam[nParamNo].Visible = true;

            g_pTxtParam[nParamNo].Text = strValue;
            g_pLabParamTitle[nParamNo].Text = strParamTitle;
            g_pLabParamUnit[nParamNo].Text = strParamUnit;
        }
        private void EnableAxis(int nAxis, bool bEnable)
        {
            switch (nAxis)
            {
                case 1:
                    CmbNode1.Enabled = bEnable;
                    GrpAxis1.Visible = bEnable;
                    break;
                case 2:
                    CmbNode2.Enabled = bEnable;
                    GrpAxis2.Visible = bEnable;
                    if (bEnable == false)
                    {
                        g_uESCNodeID[1] = 0;
                        g_uESCSlotID[1] = 0;
                    }
                    break;
                case 3:
                    CmbNode3.Enabled = bEnable;
                    GrpAxis3.Visible = bEnable;
                    if (bEnable == false)
                    {
                        g_uESCNodeID[2] = 0;
                        g_uESCSlotID[2] = 0;
                    }
                    break;
            }
        }

    }
}
