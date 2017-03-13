using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {


        public string file;
        public int ptr_step;
        public int ptr_repeat;

        private object sender;
        Form1 rent;

        public Form2()
        {
            InitializeComponent();

        }

        public Form2(Form sender)
        {
            // TODO: Complete member initialization
            this.sender = sender;

            InitializeComponent();

            rent = (Form1)sender;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            file = textBox1.Text.ToString();
            ptr_step = int.Parse(textBox2.Text);
            ptr_repeat = int.Parse(textBox3.Text);

            rent.JumpInsert(this,null,5,ptr_step,ptr_repeat);

            this.Close();
        }


    }
}
