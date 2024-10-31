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
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using ZXing;
using AForge.Video;
using AForge.Video.DirectShow;
using System.Text.RegularExpressions;

//installed Install-Package ZXing.Net
// Install-Package AForge.Video
//Install - Package AForge.Video.DirectShow



namespace 乾燥炉点検アプリ
{
    public partial class Form1 : Form
    {
        OleDbCommand command = new OleDbCommand();
        OleDbConnection cnAccess = new OleDbConnection();
        OleDbDataReader oleReader;
        bool pass = false;
        private Timer timer;
        private VideoCaptureDevice videoSource;
        private TextBox activeTextBox;
        string worker = "";
        public Form1()
        {
            InitializeComponent();
            //InitializeCamera();
            cnAccess.ConnectionString = Properties.Settings.Default.真空乾燥炉データベースConnectionString;

            textBox1.Enter += new EventHandler(textBox_Enter);
            textBox2.Enter += new EventHandler(textBox_Enter);
            
        }
        private void InitializeCamera()
        {
            StopCamera();
            var videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            videoSource = new VideoCaptureDevice(videoDevices[1].MonikerString);
            videoSource.NewFrame += new NewFrameEventHandler(video_NewFrame);
            videoSource.Start();

            timer = new Timer();
            timer.Interval = 100; // Adjust the interval as needed
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();
            
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            
            activeTextBox = sender as TextBox;  
            StopCamera();
            pictureBox1.Image = null;
            InitializeCamera();
        }
        private void timer_Tick(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                //BarcodeReader reader = new BarcodeReader();
                // Create a barcode reader instance with specific options
                var reader = new BarcodeReader
                {
                    AutoRotate = true,
                    Options = new ZXing.Common.DecodingOptions
                    {
                        TryHarder = true,
                        PossibleFormats = new[] { BarcodeFormat.QR_CODE, BarcodeFormat.CODE_128 }
                    }
                };
                var result = reader.Decode((Bitmap)pictureBox1.Image);
                if (result != null)
                {
                    activeTextBox.Text = ConvertToUtf8(result.Text);
                    StopCamera();
                }
            }
        }

        static string ConvertToUtf8(string input)
        {
            // Convert the string to a byte array using UTF-8 encoding
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(input);

            // Convert the byte array back to a string
            string utf8String = Encoding.UTF8.GetString(utf8Bytes);

            string RemoveInvisibleCharacters(string input1)
            {
                // Define a regex pattern to match invisible characters
                string pattern = @"[\u200B-\u200D\uFEFF]";
                return Regex.Replace(input1, pattern, string.Empty);
            }

            string cleanedString = RemoveInvisibleCharacters(utf8String);

            return cleanedString;
        }
        private void StopCamera()
        {
            if (videoSource != null)
            {
                if (videoSource.IsRunning)
                {
                    videoSource.SignalToStop();
                    videoSource.WaitForStop();
                }
                videoSource.NewFrame -= new NewFrameEventHandler(video_NewFrame);
                videoSource = null;
                timer?.Stop();
                timer?.Dispose();
            }
        }
            
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap bitmap = (Bitmap)eventArgs.Frame.Clone();
            pictureBox1.Image = bitmap;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            label4.Text = DateTime.Now.ToString();
            textBox1.Focus();
            if (!string.IsNullOrEmpty(Global.Uid))
            {
                textBox1.Text = Global.Uid;
            }
        }

        //日常点検
        private void button1_Click(object sender, EventArgs e)
        {
            StopCamera();
            string registerDate = "";
            login_Method();
            if (pass == true)
            {
                try
                {
                    cnAccess.Open();
                    command.Connection = cnAccess;
                    command.CommandText = "SELECT * FROM 真空乾燥炉 WHERE 乾燥炉機番 = @inputText AND 毎日日付3桁連番 LIKE @inputText1";
                    command.Parameters.AddWithValue("@inputText", textBox2.Text);
                    command.Parameters.AddWithValue("@inputText1",DateTime.Today.ToString("yyyyMMdd")+"%");
                    oleReader = command.ExecuteReader();
                    while (oleReader.Read())
                    {
                        registerDate = oleReader["日常点検日時"].ToString();
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

                if (string.IsNullOrEmpty(registerDate))
                {
                    
                    if (textBox2.Text.Substring(textBox2.Text.Length - 1) == "3")
                    {
                        this.Hide();
                        機番3_画面 gamen_Third = new 機番3_画面();
                        gamen_Third.kikaiNumber = textBox2.Text;
                        gamen_Third.worker = textBox1.Text;
                        gamen_Third.Show();
                    }
                    else if (textBox2.Text.Substring(textBox2.Text.Length - 1) == "2")
                    {
                        this.Hide();
                        機番2_画面 gamen_Second = new 機番2_画面();
                        gamen_Second.kikaiNumber = textBox2.Text;
                        gamen_Second.worker = textBox1.Text;
                        gamen_Second.Show();
                    }
                    else if (textBox2.Text.Substring(textBox2.Text.Length - 1) == "1")
                    {
                        this.Hide();
                        機番1_画面 gamen_First = new 機番1_画面();
                        gamen_First.kikaiNumber = textBox2.Text;
                        gamen_First.worker = textBox1.Text;
                        gamen_First.Show();
                    }
                }
                else
                {
                    MessageBox.Show("この機械の本日の点検は終わってます");
                    //textBox1.Clear();
                    textBox2.Clear();
                }
            }
            

        }

        //login
        private void login_Method()
        {
            //pictureBox1.Dispose();
            
            try
            {
                
                cnAccess.Open();
                command.Connection = cnAccess;
                command.CommandText = "SELECT * FROM ユーザ WHERE 社員ID ='" + textBox1.Text + "'";
                oleReader = command.ExecuteReader();
                while (oleReader.Read())
                {
                    worker = oleReader["ID"].ToString();
                }
                oleReader.Close();
            }
            catch (Exception ex)
            {
                //throw new InvalidOperationException("Something went wrong!");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                cnAccess.Close();
            }
            string kikaiBango = "";
            if (!string.IsNullOrEmpty(worker))
            {
                
                try
                {
                    cnAccess.Open();
                    command.Connection = cnAccess;
                    command.CommandText = "SELECT * FROM 乾燥炉機番情報 WHERE 機番 ='" + textBox2.Text + "'";
                    oleReader = command.ExecuteReader();
                    while (oleReader.Read())
                    {
                        kikaiBango = oleReader["ID"].ToString();
                    }
                    oleReader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    cnAccess.Close();
                }
            }
            else
            {
                MessageBox.Show("作業者を見つかりませんでした。");
            }

            //passs to form or not
            if (string.IsNullOrEmpty(kikaiBango))
            {
                MessageBox.Show("機番が見つかりませんでした。");
            }
            else
            {
                pass = true;
                string latestDate = "XXXX/XX/XX";
                try
                {
                    cnAccess.Open();
                    command.Connection = cnAccess;
                    command.CommandText = "SELECT * FROM 真空乾燥炉 WHERE 乾燥炉機番 = '" + textBox2.Text + "' ORDER BY オイル交換日時 ASC";
                    oleReader = command.ExecuteReader();
                    while (oleReader.Read())
                    {
                        latestDate = oleReader["オイル交換日時"].ToString();
                    }
                    oleReader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    cnAccess.Close();
                }
                if(latestDate == "")
                {
                    MessageBox.Show("オイル交換の登録がありません");
                }
                else if (latestDate == "XXXX/XX/XX")
                {
                    MessageBox.Show("オイル交換の登録がありません");
                }
                else
                {
                   int totalDays =(DateTime.Today - DateTime.Parse(latestDate)).Days;
                    int remainingDays = 180 - totalDays;
                    MessageBox.Show("オイル交換までの残り"+remainingDays+"日です。");
                }
            }
            
        }

        //oil change 
        private void button2_Click(object sender, EventArgs e)
        {
            login_Method();
            if(pass == true)
            {
                this.Hide();
                オイル交換画面 oilChange = new オイル交換画面();
                oilChange.kikaiNumber = textBox2.Text;
                oilChange.worker = textBox1.Text;
                oilChange.Show();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopCamera();
            //base.OnFormClosing(e);
            Application.Exit();

        }

        //worker name
        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox1.Focus();
            Global.Uid = null;
            label6.Text = null;

        }

        //kikai
        private void button5_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox2.Focus();
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                try
                {

                    cnAccess.Open();
                    command.Connection = cnAccess;
                    command.CommandText = "SELECT * FROM ユーザ WHERE 社員ID ='" + textBox1.Text + "'";
                    oleReader = command.ExecuteReader();
                    while (oleReader.Read())
                    {
                        worker = oleReader["作業者名"].ToString();
                    }
                    oleReader.Close();
                }
                catch (Exception ex)
                {
                    //throw new InvalidOperationException("Something went wrong!");
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    cnAccess.Close();
                }
                if (!string.IsNullOrEmpty(worker))
                {
                    label6.Visible = true;
                    label6.Text = worker + " 様";
                    Global.Uname = worker;
                    Global.Uid = textBox1.Text;
                }
            }
        }
    }
}
