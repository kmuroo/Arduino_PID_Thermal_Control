using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Arduino_PID_Thermal_Control
{
    public partial class Form2 : Form
    {
        //一つのフォームのインスタンスを保持するフィールド
        private static Form2 _instance;

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            //this.textBox1.Focus();
            //this.ActiveControl = this.textBox1;
            //ActiveControl = textBox1;
            textBox1.Select();
        }


        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            label19.Text = textBox1.Text;
            label20.Text = textBox2.Text;
            label21.Text = textBox3.Text;
            label22.Text = textBox4.Text;
            label23.Text = textBox5.Text;
            label24.Text = textBox6.Text;
            label25.Text = textBox7.Text;
            label26.Text = textBox8.Text;

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }


        public string Form2_Text1
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        public string Form2_Text2
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }
        public string Form2_Text3
        {
            get { return textBox3.Text; }
            set { textBox3.Text = value; }
        }
        public string Form2_Text4
        {
            get { return textBox4.Text; }
            set { textBox4.Text = value; }
        }
        public string Form2_Text5
        {
            get { return textBox5.Text; }
            set { textBox5.Text = value; }
        }
        public string Form2_Text6
        {
            get { return textBox6.Text; }
            set { textBox6.Text = value; }
        }
        public string Form2_Text7
        {
            get { return textBox7.Text; }
            set { textBox7.Text = value; }
        }
        public string Form2_Text8
        {
            get { return textBox8.Text; }
            set { textBox8.Text = value; }
        }

        public static Form2 Instance
        {
            get
            {
                //_instanceがnullまたは破棄されているときは、
                //新しくインスタンスを作成する
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new Form2();
                }
                return _instance;
            }
        }
        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }
    }
}
