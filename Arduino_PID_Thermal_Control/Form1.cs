using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
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
                    textBox1.AppendText("PID thermal controllerをオープンしています\n\r\n\r");
                                       
                    serialPort1.ReadExisting(); //バッファを空に

                    if (devicecheck() == 0) //PID thermal controllerかどうかのチェック
                    {
                        textBox1.AppendText("PID thermal controller を接続しました (" + serialPort1.PortName + ")\n\r\n\r");
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
                textBox1.AppendText("device チェック終了\n\r\n\r");
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
                textBox1.AppendText("COMポートからが反応がありません\n\r\n\r");
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
                textBox1.AppendText("PID thermal controller を切断しました (" + serialPort1.PortName + ")\n\r\n\r");
                button4.BackColor = Color.LightGray;
                button1.BackColor = SystemColors.Control;
                button3.BackColor = Color.LightGray;
                button2.BackColor = SystemColors.Control;
            }
        }

        private string arduino_send_recv(string send_message)//Arduinoにメッセージ送信、コールバックあり
        {
            string recv_message;

            serialPort1.ReadExisting(); //バッファを空に
            serialPort1.Write(send_message); //メッセージ送信

            serialPort1.ReadTimeout = 5000;
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
                    serialPort1.ReadExisting(); //バッファを空に
                    arduino_send("e");
                    textBox1.AppendText("Monitor start\n\r\n\r");
                    button2.BackColor = Color.LightGray;
                    button3.BackColor = SystemColors.Control;
                }
                else
                {
                    textBox1.AppendText("COMポ－トが開かれていません\n\r\n\r");
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
                textBox1.AppendText(serialPort1.ReadLine() + "\n\r");
            }
            arduino_send("s");
            monitor_enable = false;
            textBox1.AppendText("Monitor stop\n\r\n\r");
            return 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (monitor_enable == true)
            {
                monitor_enable = false;
                if (serialPort1.IsOpen)
                {
                    button3.BackColor = Color.LightGray;
                    button2.BackColor = SystemColors.Control;
                }
                else
                {
                    textBox1.AppendText("COMポートが開かれていません\n\r\n\r");
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
                    serialPort1.Close();
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
                MessageBox.Show("Controller RESET\n\r\n\r");
                monitor_enable = false;
                comclose();
            }
            else
            {
                textBox1.AppendText("ポートは開かれていません\n\r\n\r");
            }
        }
    }
}
