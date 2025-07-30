using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = System.Windows.Forms.Application;


namespace Arduino_PID_Thermal_Control
{
    public partial class Form1 : Form
    {
        string default_portname = "COM7"; // デフォルトCOMポート
        string[] ports;
        bool monitor_enable = false;
        bool verbose = false;// 詳細表示フラグ
        string[] current_parameters;
        string temp_file_name;
        StreamWriter temp_file;
        int job = 0;
        int iteration_number = 0;

        public Form1()
        {
            InitializeComponent();
            FormClosing += Form1_FormClosing;
            add_serial_portname();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen == true)
            {
                comclose();
            }
            else
            {
                button4.BackColor = Color.LightGray;
                button1.BackColor = SystemColors.Control;
            }
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;
        }

        private void add_serial_portname()
        {
            ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            comboBox1.SelectedIndex = comboBox1.Items.Count - 1;

            if (comboBox1.FindString(default_portname) > 0)
            {
                comboBox1.SelectedIndex = comboBox1.FindString(default_portname);
            }
            serialPort1.PortName = comboBox1.SelectedItem.ToString();
        }
       
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private bool comopen()
        {
            if (serialPort1.IsOpen == false)
            {
                try
                {
                    serialPort1.BaudRate = 115200; // ArduinoソースのSerial　ボーレートと合わせる
                    serialPort1.Open();
                    while (serialPort1.IsOpen == false)
                    {
                        // ポートがオープンするまで待つ
                    }
                    textBox1.AppendText("PID thermal controllerをオープンしています\r\n\r\n");
                                       
                    serialPort1.ReadExisting(); //バッファを空に

                    if (devicecheck() == 0) //PID thermal controllerかどうかのチェック
                    {
                        textBox1.AppendText("PID thermal controller を接続しました (" + serialPort1.PortName + ")\r\n\r\n");
                        current_parameters = get_current_parameters();
                        verbose = false; // デフォルトでは詳細表示をオフにする
                        label9.Text = current_parameters[0];
                        label10.Text = current_parameters[1];
                        label11.Text = current_parameters[2];
                        label12.Text = current_parameters[3];
                        label13.Text = current_parameters[4];
                        label14.Text = current_parameters[5];
                        label15.Text = current_parameters[6];
                        label16.Text = current_parameters[7];
                        return true;
                    }
                    else {
                        textBox1.AppendText("誤ったCOMポートを選択しています。他のCOMポートを試してください。\r\n\r\n");
                        comclose();
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    textBox1.AppendText(ex.Message);
                    textBox1.AppendText("COMポートオープンに失敗しました。別のCOMポートを試してください。\r\n\r\n");
                    comclose();
                    return false;
                }
            }
            return true;
        }


        private int devicecheck()
        {
            try
            {
                textBox1.AppendText("device チェック開始\n\r\n\r");
                string r = arduino_send_recv("r");
                serialPort1.ReadExisting(); //バッファを空に
                string a = "PID_Controller";
                string c = arduino_send_recv("c");
                c = c.TrimEnd('\r', '\n'); //改行コード削除
                textBox1.AppendText("device チェック終了\r\n\r\n");
                if (a == c) //接続成功すれば "PID_Controller" がArduinoから返ってくる
                {
                    return 0; //正しければ0を返す
                }
                else
                {
                    return 1; //間違っていれば1を返す
                }
            }
            catch (Exception ex)
            {
                textBox1.AppendText(ex.Message);
                textBox1.AppendText("COMポートからが反応がありません\r\n\r\n");
                comclose();
                return -1;
            }
        }



        private void comclose()
        {

            if (serialPort1.IsOpen == true)
            {
                arduino_send("r");
                serialPort1.Close();
                textBox1.AppendText("PID thermal controller を切断しました (" + serialPort1.PortName + ")\r\n\r\n");
                button4.BackColor = Color.LightGray;
                button1.BackColor = SystemColors.Control;
                button3.BackColor = Color.LightGray;
                button2.BackColor = SystemColors.Control;
            }
        }

        private string[] get_current_parameters() //現在のパラメーター収得
        {
            string[] current_parameters;
            string parameters;
            serialPort1.ReadExisting(); //バッファを空に
            parameters = arduino_send_recv("l");
            current_parameters = parameters.Split(','); //カンマ区切りで分割
            return current_parameters;
        }

        private string arduino_send_recv(string send_message)//Arduinoにメッセージ送信、コールバックあり
        {
            string recv_message;
            serialPort1.ReadExisting(); //バッファを空に
            serialPort1.Write(send_message); //メッセージ送信
            serialPort1.ReadTimeout = 5000; //タイムアウトを5秒に設定
            recv_message = serialPort1.ReadLine();//メッセージ受信
            return recv_message;
        }

        private void arduino_send(string send_message)//Arduinoにメッセージ送信、コールバックなし
        {
            serialPort1.Write(send_message);
        }

 
        private void button1_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                serialPort1.PortName = comboBox1.SelectedItem.ToString();
                if (comopen() == true)
                {
                    button1.BackColor = Color.LightGray;
                    button4.BackColor = SystemColors.Control;
                    button6.Enabled = true;
                    button7.Enabled = true;
                    button9.Enabled = true;
                    button10.Enabled = true;
                    //button8.Enabled = true;
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (monitor_enable == false)
            {
                monitor_enable = true;
                if (serialPort1.IsOpen)
                {
                    job++;
                    serialPort1.ReadExisting(); //バッファを空に
                    arduino_send("e");
                    textBox1.AppendText("Monitor start  Job. " + job + "\r\n");
                    button2.BackColor = Color.LightGray;
                    button3.BackColor = SystemColors.Control;
                    button7.Enabled = false;
                    button8.Enabled = false;
                    button9.Enabled = false;
                    button10.Enabled = false;
                    iteration_number = 0;
                   
                    if(File.Exists(temp_file_name))
                    {
                        temp_file.Close();
                        File.Delete(temp_file_name);
                    }
                    temp_file_name = Path.GetTempFileName();
                    Encoding enc = Encoding.GetEncoding("Shift_JIS");
                    temp_file = new StreamWriter(temp_file_name);
                    DateTime dateTime = DateTime.Now;
                    temp_file.WriteLine("#" + dateTime.ToString("yyyy/MM/dd  HH:mm:ss") + "  Job. " + job.ToString());
                }
                else
                {
                    textBox1.AppendText("COMポ－トが開かれていません\r\n\r\n");
                }
                
                Task<int> task = Task.Run(() => {
                    return report_run();
                });
                
            }
        }

        private int report_run()
        {
            while (monitor_enable)
            {
                iteration_number++;
                string report = serialPort1.ReadLine(); //Arduinoからのデータを読み込む
                report = report.Replace("\r",""); // 改行コード\rを削除
                report = report.Replace("\n", ""); // 改行コード\nを削除
                if (verbose)
                {
                    if (iteration_number % 3 == 1)
                    {
                        int i = iteration_number / 3 + 1;
                        textBox1.AppendText(i.ToString() + "\t" + report + "\r\n");
                        temp_file.Write(i.ToString() + "," + report.Replace("\t", ",") + ",");
                    }
                    else if(iteration_number % 3 == 2)
                    {
                        textBox1.AppendText(report + "\r\n");
                        temp_file.Write(report.Replace("\t", ",") + ",");
                    }
                    else// if (iteration_number % 3 == 0)
                    {
                        textBox1.AppendText(report + "\r\n");
                        temp_file.WriteLine(report.Replace("\t", ","));
                    }
                }
                else
                {
                    textBox1.AppendText(iteration_number.ToString() + "\t" + report + "\r\n");
                    temp_file.WriteLine(iteration_number.ToString() + "," + report.Replace("\t", ","));
                }
            }
            arduino_send("s");
            monitor_enable = false;
            textBox1.AppendText("Monitor stop\r\n\r\n");
            return 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (monitor_enable == true)
            {
                monitor_enable = false;
                button8.Enabled = true;
                button7.Enabled = true;
                button9.Enabled = true;
                button10.Enabled = true;
                if (serialPort1.IsOpen)
                {
                    button3.BackColor = Color.LightGray;
                    button2.BackColor = SystemColors.Control;
                }
                else
                {
                    textBox1.AppendText("COMポートが開かれていません\r\n\r\n");
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                if (monitor_enable == false)
                {
                    comclose();
                    button7.Enabled = false;
                    button9.Enabled = false;
                    button10.Enabled = false;
                    verbose = false; // 詳細表示をオフにする
                    button9.BackColor = SystemColors.Control;
                    //button8.Enabled = false;
                }
                else
                {
                    MessageBox.Show("Controller is now running");
                }
            }
        }



        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 質問ダイアログを表示する
            DialogResult result = MessageBox.Show("終了しますか？", "質問", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                // はいボタンをクリックしたときはウィンドウを閉じる
                e.Cancel = true;
            }
            else
            {
                //COMポートを閉じて終了
                
                if (serialPort1.IsOpen)
                {
                    arduino_send("r");
                    serialPort1.Close();
                }

                if (temp_file != null)
                {
                    temp_file.Close();
                    File.Delete(temp_file_name);
                }
                
                
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                arduino_send("r");
                MessageBox.Show("Controller RESET\r\n\r\n");
                monitor_enable = false;
                button7.Enabled = false;
                button9.Enabled = false;
                button10.Enabled = false;
                comclose();
            }
            else
            {
                textBox1.AppendText("ポートは開かれていません\r\n\r\n");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (monitor_enable == true)
            {
                MessageBox.Show("Monitor中です。Monitorを中止してください\r\n\r\n");
                return;
            }
            if (serialPort1.IsOpen == false)
            {
                textBox1.AppendText("COMポートが開かれていません\r\n\r\n");
                return;
            }
            Form2.Instance.Form2_Text1 = label9.Text;
            Form2.Instance.Form2_Text2 = label10.Text;
            Form2.Instance.Form2_Text3 = label11.Text;
            Form2.Instance.Form2_Text4 = label12.Text;
            Form2.Instance.Form2_Text5 = label13.Text;
            Form2.Instance.Form2_Text6 = label14.Text;
            Form2.Instance.Form2_Text7 = label15.Text;
            Form2.Instance.Form2_Text8 = label16.Text;
            Form2.Instance.ShowDialog();
            if (Form2.Instance.DialogResult == DialogResult.OK)
            {
                // OKボタンがクリックされた場合、パラメータを更新
                label9.Text = Form2.Instance.Form2_Text1;
                label10.Text = Form2.Instance.Form2_Text2;
                label11.Text = Form2.Instance.Form2_Text3;
                label12.Text = Form2.Instance.Form2_Text4;
                label13.Text = Form2.Instance.Form2_Text5;
                label14.Text = Form2.Instance.Form2_Text6;
                label15.Text = Form2.Instance.Form2_Text7;
                label16.Text = Form2.Instance.Form2_Text8;
                serialPort1.ReadExisting();

                string response = arduino_send_recv("p" + label9.Text + "," + label10.Text + "," + label11.Text
                    + "," + label12.Text + "," + label13.Text + "," + label14.Text + ","
                    + label15.Text + "," + label16.Text);
                response = response.TrimEnd('\r', '\n');

                if ( response == "P")
                {
                    textBox1.AppendText("Update parameters.\n\r\n\r");
                }
                else
                {
                    textBox1.AppendText("Failed to update parameters.\n\r\n\r");
                }

                serialPort1.ReadExisting();
                current_parameters = get_current_parameters();
                textBox1.AppendText("New parameters are\n\r\n\r");
                textBox1.AppendText(current_parameters[0] + "\t" + current_parameters[1] + "\t"
                    + current_parameters[2] + "\t" +current_parameters[3] + "\n\r\n\r" + current_parameters[4] + "\t"
                    + current_parameters[5] + "\t" + current_parameters[6] + "\t" + current_parameters[7] + "\t" + "\n\r\n\r");
            }
            else if (Form2.Instance.DialogResult == DialogResult.Cancel)
            {
                // Cancelボタンがクリックされた場合、何もしない
            }
            else
            {
                MessageBox.Show("パラメータの更新に失敗しました。");

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = "data" + job.ToString() + ".csv";
            saveFileDialog1.Filter = "csv型式ファイル(*.csv)|*.csv";
            saveFileDialog1.Title = "Save an DATA File";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                try
                {
                    // 一次ファイルをsaveFileDialog1.FileNameにコピー

                    if (temp_file != null)
                    {
                        temp_file.Close();
                        File.Copy(@temp_file_name, @saveFileDialog1.FileName,true);
                        MessageBox.Show("データを保存しました: " + saveFileDialog1.FileName);
                    // 一時ファイルを削除
                        File.Delete(temp_file_name);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("データの保存に失敗しました: " + ex.Message);
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                if (verbose == true)
                {

                    string t = arduino_send_recv("t");
                    verbose = false;
                    button9.BackColor = SystemColors.Control;
                    button9.Text = "Terse Mode";
                    textBox1.AppendText("簡易レポートモードに変更\r\n\r\n");
                }
                else
                {
                    string v = arduino_send_recv("v");
                    verbose = true;
                    button9.BackColor = Color.LightGray;
                    button9.Text = "Verbose Mode";
                    textBox1.AppendText("冗長レポートモードに変更\r\n\r\n");
                }
            }
        }

     
        private void button10_Click(object sender, EventArgs e)
        {
            button10.BackColor = Color.LightGray;
            string I = arduino_send_recv("i");
            button10.BackColor = SystemColors.Control;
        }
    }
}
