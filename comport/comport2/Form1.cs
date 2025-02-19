using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace comport2
{

    public partial class Form1 : Form
    {
        private SerialPort serialPort;
        private bool isComboBox5DataSent = false;

        #region 
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private LowLevelKeyboardProc _proc;
        private static IntPtr _hookID = IntPtr.Zero;
        private string scanner_data = string.Empty;
        #endregion
        public Form1()
        {
            _proc = HookCallback;
            _hookID = SetHook(_proc);
            InitializeComponent();
            textBox3.TextChanged += TextBox3_TextChanged;
        }
        ~Form1()
        {
            UnhookWindowsHookEx(_hookID);
        }
        #region read input data
        
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }


        public bool system_on = false;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        public string ini;
        public int barcode_length ;
        int m = 0;
        int count = 0;
        public bool flag1 = false;
        bool flag2 = false;
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (flag1)
            {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
                {
                    
                    int vkCode = Marshal.ReadInt32(lParam);
                    Keys key = (Keys)vkCode;
                    string k = key.ToString();
                    if (m < barcode_length)
                    {
                        if (k.Length == 2 )
                        {
                            k = k[1].ToString();
                        }
                        if (k.Length < 3)
                        {
                            if (m < ini.Length)
                            {

                                if (k == ini[m].ToString())
                                {
                                    scanner_data += k;
                                    m++;

                                }
                                else
                                {
                                    m = 0;
                                    scanner_data = "";
                                }
                            }
                            else
                            {
                                scanner_data += k;
                                m++;
                            }
                        }
                    }
                    if (m == barcode_length && k == "Return")
                    {
                        if ( flag2 || brdchk(scanner_data) )
                        {
                            this.Controls.Add(textBox4);
                            textBox4.Text = scanner_data;
                            add2xml(scanner_data);
                            count++;
                            textBox3.Text = count.ToString();
                            m = 0;
                            scanner_data = "";
                            SendPulse("1");
                        }
                        
                    }
                    if (k == "Return")
                    {
                        m = 0;
                        scanner_data = "";
                        if (!system_on)
                        {
                            Thread.Sleep(100);
                            this.button1.PerformClick();

                        }

                    }
                }
            }
            else
            {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    Keys key = (Keys)vkCode;
                    string k = key.ToString();
                    if(k == "Return")
                    {
                        button1_Click(null, null);
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
        #endregion 
        string initial_barcode = "";



        public bool brdchk(string brcode)
        {
          //  return true;
                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "sampleqwerty.xml");
                var Data = new Dictionary<string, object>();
                if (File.Exists(filePath))
                {
                    XElement root = XElement.Load(filePath);
                    foreach (var item in root.Elements())
                    {
                       if(brcode == item.Value)
                        {
                            return false;
                        }
                    }
                    return true;
                }
                else
                {
                    return true;
                }
            }
        public void add2xml(string brcode)
        {
            string tmp;
                try
            {
                    string filePath = Path.Combine(Directory.GetCurrentDirectory(), "sampleqwerty.xml");
                    var Data = new Dictionary<string, object>();
                    if (File.Exists(filePath))
                    {
                        XElement root = XElement.Load(filePath);
                        foreach (var item in root.Elements())
                        {
                            Data[item.Name.LocalName] = item.Value;
                        }
                    }
                try
                {
                   tmp = Data["key"].ToString();
                }
                catch {
                    Data["key"] = 0;
                    tmp = "0";
                }
                int c = int.Parse(tmp);
                if (c > 1000)
                {
                    c = -1;
                }
                string cnt = "key" + (c+1).ToString();
                Data[cnt] = brcode;
                Data["key"] = (c+1).ToString();
                    
                    XElement xml = new XElement("Root", new List<XElement>(Data.Select(kvp => new XElement(kvp.Key, kvp.Value))));
                    xml.Save(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occurred: " + ex.Message);
                }
            
        }
        public void delay(int k)
        {
            DateTime x = DateTime.Now;
            while((DateTime.Now - x).TotalSeconds < k) {}
        }
        private bool SendPulse( String a)
        {
           
            try
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();
                }
                serialPort1.Open();
                serialPort1.Write(a);
                return true;
            }
            catch (Exception ex)
            {
                cnt = false;
                Console.WriteLine(ex.Message);
                IntPtr hWnd = FindWindow(null, "COM Error");
                if (hWnd == IntPtr.Zero)
                {
                    MessageBox.Show($"Error : {ex.Message}", "COM Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
                

            }
            finally
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();
                    
                }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // Thread.Sleep(2000);
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "sampleqwerty.xml");
            var Data = new Dictionary<string, object>();
            if (File.Exists(filePath))
            {
                XElement root = XElement.Load(filePath);
                foreach (var item in root.Elements())
                {
                    Data[item.Name.LocalName] = item.Value;
                }
            }
            comboBox1.SelectedItem = Data["COM"];
            comboBox2.SelectedItem = Data["bd"];
            comboBox3.SelectedItem = Data["bit"];
            comboBox4.SelectedItem = Data["stopbit"];
          //  Thread.Sleep(1000);
            this.MaximizeBox = false;
            this.textBox3.ReadOnly = true;
         //   this.textBox4.ReadOnly = true;
            string none = loaddata2dd(10008);
            comboBox5.SelectedIndex = (int.Parse(none)-1);
            

        }
        public string loaddata2dd(int k)
        {
            string data;
            Thread.Sleep(1000);
            Excel.Application excelApp = new Excel.Application();
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "model.xlsx");

            try
            {
                if (File.Exists(filePath))
                {
                    string str;
                    Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                    Excel.Worksheet worksheet = workbook.Worksheets[1];

                    if (k == 10008)
                    {
                        for (int row = 1; row <= worksheet.UsedRange.Rows.Count; row++)
                        {
                            string c1 = (worksheet.Cells[row, 1] as Excel.Range)?.Value2?.ToString();
                            comboBox5.Items.Add(c1);
                        }

                        // Return value from cell D100
                        str = worksheet.Cells[1, 4]?.Value2?.ToString();
                    }
                    else if (k > 1000)
                    {
                        str = (worksheet.Cells[k - 1000, 2] as Excel.Range)?.Value2?.ToString();
                        //return c1;
                    }
                    else
                    {

                        str = (worksheet.Cells[k, 3] as Excel.Range)?.Value2?.ToString();
                        // Write value of k to cell D100
                        worksheet.Cells[1, 4] = k.ToString();

                    }

                    workbook.Save();
                    workbook.Close(false);

                    // Cleanup
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                    return str;
                }
                else
                {
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = workbook.Worksheets[1];

                    // Write value of k to cell D100
                    worksheet.Cells[1, 4] = k.ToString();

                    // Save new workbook
                    workbook.Save();
                    workbook.Close(false);

                    // Cleanup
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                Console.WriteLine("Error: " + ex.Message);
            }

            return "none"; // Default return value
        }
        public void portchange()
        {
            
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "sampleqwerty.xml");
            var Data = new Dictionary<string, object>();
            if (File.Exists(filePath))
            {
                XElement root = XElement.Load(filePath);
                foreach (var item in root.Elements())
                {
                    Data[item.Name.LocalName] = item.Value;
                }
            }
            Data["COM"] = comboBox1.SelectedItem;
            Data["bd"] = comboBox2.SelectedItem ;
            Data["bit"] = comboBox3.SelectedItem;
            Data["stopbit"] = comboBox4.SelectedItem;
            serialPort1.PortName = "COM" + Data["COM"].ToString();
            serialPort1.BaudRate = int.Parse(comboBox2.SelectedItem.ToString()) ;
            serialPort1.DataBits = int.Parse(comboBox3.SelectedItem.ToString());
            serialPort1.Parity = Parity.Odd;
     //       serialPort.StopBits = StopBits.One;


            XElement xml = new XElement("Root", new List<XElement>(Data.Select(kvp => new XElement(kvp.Key, kvp.Value))));
            xml.Save(filePath);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        public bool cnt = true;
      //  Label error = new Label();//
        private void button1_Click(object sender, EventArgs e)
        {

            label8.Text = "";
            bool ck = brdchk("none");
            portchange();
            if(textBox4.Text == "global")
            {
                textBox4.Text = "";
                this.textBox4.ReadOnly = true;
                flag2 = true;
            }
            int si = comboBox5.SelectedIndex;
            if (si < 0)
            {
                si = 0;
            }
            comboBox5.SelectedIndex = si;
           // SendPulse("1")
            if ( true)
            {
                if (si != 0)
            {
                textBox1.Text = loaddata2dd(1001 + si);
                textBox2.Text = (loaddata2dd(si + 1).ToString());
            }
          
                if (textBox1.Text.Length > 0 && textBox2.Text.Length > 0)
                {
                    ini = (textBox1.Text);
                    ini = ini.ToUpper();
                    barcode_length = int.Parse(textBox2.Text);
                    flag1 = true;
                    textBox3.Text = 0.ToString();
                    //foreach (Control control in groupBox1.Controls)
                    //{
                    //    control.Enabled = false;
                    //}
                    foreach (Control control in groupBox2.Controls)
                    {
                        control.Enabled = false;
                    }
                    foreach (Control control in groupBox1.Controls)
                    {
                        control.Enabled = false;
                    }

                    button1.BackColor = Color.LimeGreen;

                }
                else
                {

                    MessageBox.Show($"Error : Enter valid Parameter", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            string textBox3Data = textBox3.Text;
            string textBox5Data = textBox5.Text;

            // Get comboBox5 data (send only once)
            string comboBox5Data = isComboBox5DataSent || comboBox5.SelectedItem == null
                ? string.Empty
                : comboBox5.SelectedItem.ToString();

            // Prepare the message in the format "TextBox5: <value>, ComboBox5: <value>, TextBox3: <value>"
            string message = $"{textBox5Data},{comboBox5Data},{textBox3Data}";

            // Send message only once for ComboBox5 data
            //if (!string.IsNullOrEmpty(comboBox5Data))
            //{
            //    isComboBox5DataSent = true; // Mark comboBox5 data as sent
            //}

            // Replace with the public or local IP address of the receiver
            string receiverIp = textBox6.Text;  // You can replace this with your receiver's IP address
            int receiverPort = 5000;  // Receiver listens on this port

            try
            {
                using (TcpClient client = new TcpClient(receiverIp, receiverPort))
                using (NetworkStream stream = client.GetStream())
                {
                    byte[] data = Encoding.UTF8.GetBytes(message);
                    stream.Write(data, 0, data.Length);
                }

                Console.WriteLine("Data sent: " + message); // Log the sent data for debugging
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to send data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
//NIRMAL
//RITHVI
//240711