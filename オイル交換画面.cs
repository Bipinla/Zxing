using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;

namespace 乾燥炉点検アプリ
{
    public partial class オイル交換画面 : Form
    {
        OleDbCommand command = new OleDbCommand();
        OleDbConnection cnAccess = new OleDbConnection();
        OleDbDataReader oleReader;
        private string latestDate = "XXXX/XX/XX";
        private bool isTabTipStarted = false;
        public string kikaiNumber { get; set; }
        public string worker { get; set; }
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        //private const uint WM_SYSCOMMAND = 0x0112;
        //private const uint SC_CLOSE = 0xF060;

        private const uint WM_INPUTLANGCHANGEREQUEST = 0x0050;
        private static readonly IntPtr INPUTLANG_JAPANESE = new IntPtr(0x0411);
        //private static readonly IntPtr INPUTLANG_ENGLISH = new IntPtr(0x0409); // US English
        public オイル交換画面()
        {
            InitializeComponent();
            cnAccess.ConnectionString = Properties.Settings.Default.真空乾燥炉データベースConnectionString;
            textBox1.GotFocus += TextBox1_GotFocus;
            textBox1.Leave += TextBox1_LeaveFocus;
        }

        private void TextBox1_GotFocus(object sender, EventArgs e)
        {
            string oskPath = Path.Combine(Application.StartupPath, "TabTip.exe");
            if (!isTabTipStarted)
            {
                var tabTipProcesses = Process.GetProcessesByName("TabTip");
                foreach (var process in tabTipProcesses)
                {
                    process.Kill();
                    process.WaitForExit();
                }

                // Set the input language to Japanese
                IntPtr hWnd = GetForegroundWindow();
                PostMessage(hWnd, WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, INPUTLANG_JAPANESE);

                // Start a new TabTip process
                Process.Start(oskPath);
            }

        }

        private void TextBox1_LeaveFocus(object sender, EventArgs e)
        {

        }

        private void オイル交換画面_Load(object sender, EventArgs e)
        {
            label7.Text = Global.Uname + "様";
            WindowState = FormWindowState.Maximized;
            label2.Text = DateTime.Now.ToString();
            try
            {
                cnAccess.Open();
                command.Connection = cnAccess;
                command.CommandText = "SELECT * FROM 真空乾燥炉 WHERE 乾燥炉機番 = '"+kikaiNumber+ "' ORDER BY オイル交換日時 ASC";
                oleReader = command.ExecuteReader();
                while (oleReader.Read())
                {
                    latestDate = oleReader["オイル交換日時"].ToString();
                }
                oleReader.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                cnAccess.Close();
            }
            if(latestDate == "")
            {
                label5.Text = latestDate;
            }
            else
            {
                label5.Text = DateTime.Parse(latestDate).ToShortDateString();
            }
            
        }

        //Home 
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 homePage = new Form1();
            homePage.Show();
        }

        //register
        private void button1_Click(object sender, EventArgs e)
        {
            string seijyobango = "";
            string newSeijyobango = "";
            try
            {
                cnAccess.Open();
                command.Connection = cnAccess;
                command.CommandText = "SELECT TOP 1 毎日日付3桁連番 FROM 真空乾燥炉 WHERE 毎日日付3桁連番 LIKE '" + DateTime.Today.ToString("yyyyMMdd") + "%' ORDER BY 毎日日付3桁連番 DESC";
                oleReader = command.ExecuteReader();
                while (oleReader.Read())
                {
                    seijyobango = oleReader["毎日日付3桁連番"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                cnAccess.Close();
            }

            if (string.IsNullOrEmpty(seijyobango))
            {
                newSeijyobango = DateTime.Today.ToString("yyyyMMdd") + "001";
            }
            else
            {
                if (int.Parse(seijyobango.Substring(seijyobango.Length - 3)) < 9)
                {
                    newSeijyobango = DateTime.Today.ToString("yyyyMMdd") + "00" + (int.Parse(seijyobango.Substring(seijyobango.Length - 3)) + 1).ToString();
                }
                else if (int.Parse(seijyobango.Substring(seijyobango.Length - 3)) < 99)
                {
                    newSeijyobango = DateTime.Today.ToString("yyyyMMdd") + "0" + (int.Parse(seijyobango.Substring(seijyobango.Length - 3)) + 1).ToString();
                }
                else
                {
                    newSeijyobango = DateTime.Today.ToString("yyyyMMdd") + (int.Parse(seijyobango.Substring(seijyobango.Length - 3)) + 1).ToString();
                }
            }

            try
            {
                cnAccess.Open();
                command.Connection = cnAccess;
                command.CommandText = "INSERT INTO 真空乾燥炉(毎日日付3桁連番, 乾燥炉機番, 点検作業者, 日常点検日時, オイル交換日時, 異常内容) VALUES (@newSeijyobango, @kikaiNumber, @worker, @inspectionDate, @oilChangeDate, @abnormalContent)";
                command.Parameters.AddWithValue("@newSeijyobango", newSeijyobango);
                command.Parameters.AddWithValue("@kikaiNumber", kikaiNumber);
                command.Parameters.AddWithValue("@worker", Global.Uname);
                command.Parameters.AddWithValue("@inspectionDate", DateTime.Now.ToString());
                command.Parameters.AddWithValue("@oilChangeDate", DateTime.Now.ToString());
                
                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    command.Parameters.AddWithValue("@abnormalContent", "記録なし");
                }
                else
                {
                    command.Parameters.AddWithValue("@abnormalContent", textBox1.Text);
                }
                command.ExecuteNonQuery();
                MessageBox.Show("入力完了しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                cnAccess.Close();
            }
            button2.PerformClick();
        }
    }
}
