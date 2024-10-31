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

namespace 乾燥炉点検アプリ
{
    public partial class 機番3_画面 : Form
    {
        OleDbCommand command = new OleDbCommand();
        OleDbConnection cnAccess = new OleDbConnection();
        OleDbDataReader oleReader;
        ComboBox[] combobox = new ComboBox[7];
        private bool passed = false;

        public string kikaiNumber { get; set; }
        public string worker { get; set; }
        public 機番3_画面()
        {
            InitializeComponent();
            cnAccess.ConnectionString = Properties.Settings.Default.真空乾燥炉データベースConnectionString;
            combobox[1] = comboBox1;
            combobox[2] = comboBox2;
            combobox[3] = comboBox3;
            combobox[4] = comboBox4;
            combobox[5] = comboBox5;
            combobox[6] = comboBox6;

        }

        private void 機番3_画面_Load(object sender, EventArgs e)
        {
            label15.Text = Global.Uname + " 様";
            WindowState = FormWindowState.Maximized;
            textBox1.Enabled = false;
            label13.Text = DateTime.Now.ToString();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox5.SelectedItem.ToString() == "異常あり")
            {
                textBox1.Enabled = true;
            }
            else
            {
                textBox1.Enabled = false;
                textBox1.Clear();
            }
            
        }

        //点検終了
        private void button1_Click(object sender, EventArgs e)
        {
            string seijyobango = "";
            string newSeijyobango = "";
            if (textBox1.Enabled == true && string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("テキストボックスに入力してください。");
            }
            else
            {
                for (int i=1; i<7; i++)
                {
                    if(combobox[i].SelectedItem == null)
                    {
                        passed = false;
                    }
                    else
                    {
                        passed = true;
                    }
                }

                if (passed == false)
                {
                    MessageBox.Show("全てのチェックしてください。");
                }
                else
                {
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


                    //input in database
                    try
                    {
                        cnAccess.Open();
                        command.Connection = cnAccess;
                        command.CommandText = "INSERT INTO 真空乾燥炉 (";
                        command.CommandText += "毎日日付3桁連番 , 乾燥炉機番 , 点検作業者 , 日常点検日時 , 点検項目1_オイル量 , 点検項目2_ポンプ音 , 点検項目3_昇温 , 点検項目4_真空引き , 点検項目5_シリンダー動作 , 異常有無 ";
                        if (textBox1.Enabled == true)
                        {
                            command.CommandText += " , 異常内容 ";
                        }
                        command.CommandText += ")";
                        command.CommandText += " VALUES ( ";
                        command.CommandText += " '" + newSeijyobango + "' , ";
                        command.CommandText += " '" + kikaiNumber + "' , ";
                        command.CommandText += " '" + Global.Uname + "' , ";
                        command.CommandText += " '" + DateTime.Now.ToString() + "' , ";
                        //オイル量
                        if (comboBox1.SelectedItem.ToString() == "異常なし")
                        {
                            command.CommandText += " False , ";
                        }
                        else
                        {
                            command.CommandText += " True , ";
                        }
                        //ポンプ音
                        if (comboBox2.SelectedItem.ToString() == "異常なし")
                        {
                            command.CommandText += " False , ";
                        }
                        else
                        {
                            command.CommandText += " True , ";
                        }
                        //昇温
                        if (comboBox3.SelectedItem.ToString() == "異常なし")
                        {
                            command.CommandText += " False , ";
                        }
                        else
                        {
                            command.CommandText += " True , ";
                        }
                        //点検項目4_真空引き
                        if (comboBox4.SelectedItem.ToString() == "異常なし")
                        {
                            command.CommandText += " False , ";
                        }
                        else
                        {
                            command.CommandText += " True , ";
                        }
                        //点検項目5_シリンダー動作 後で追加したから6
                        if (comboBox6.SelectedItem.ToString() == "異常なし")
                        {
                            command.CommandText += " False , ";
                        }
                        else
                        {
                            command.CommandText += " True , ";
                        }
                        //異常有無
                        if (comboBox5.SelectedItem.ToString() == "異常なし")
                        {
                            command.CommandText += " False  ";
                        }
                        else
                        {
                            command.CommandText += " True  ";
                        }
                        if (textBox1.Enabled == true)
                        {
                            command.CommandText += ", '" + textBox1.Text + "' ";
                        }
                        command.CommandText += ")";
                        command.ExecuteNonQuery();
                        MessageBox.Show("入力完了しました。");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Catch while uploading in database" + ex.Message);
                    }
                    finally
                    {
                        cnAccess.Close();
                    }
                    this.Hide();
                    Form1 register = new Form1();
                    register.Show();
                }
            }
        }

        private void 機番3_画面_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
