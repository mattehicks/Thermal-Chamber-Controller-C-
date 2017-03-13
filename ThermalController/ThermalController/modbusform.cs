using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;

namespace WindowsFormsApplication1
{
    public partial class modbusForm : Form
    {

        public SerialPort spxt;


        public modbusForm(SerialPort spx)
        {
            spxt = spx;

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region open serial port
            if (spxt.IsOpen == false)
            {
                try
                {
                    spxt.Open();
                   // errl.WriteLine("using COM1" + Environment.NewLine);
                }
                catch (Exception)
                {
                    try
                    {
                        spxt.Open();
                     //   errl.WriteLine("using COM2" + Environment.NewLine);
                    }
                    catch (Exception)
                    {
                        try
                        {
                            spxt.Open();
                          //  errl.WriteLine("using COM3" + Environment.NewLine);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                spxt = new SerialPort("COM4", 9600);
                                spxt.Open();
                               // errl.WriteLine("using COM4" + Environment.NewLine);
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    spxt = new SerialPort("COM6", 9600);
                                    spxt.Open();
                                   // errl.WriteLine("using COM6" + Environment.NewLine);
                                }
                                catch (Exception)
                                {
                                   // errl.WriteLine("Comports not ready" + Environment.NewLine);
                                    MessageBox.Show("PC COMs not ready", "MsgBox",
                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);

                                }
                            }
                        }
                    }
                }
            }

            else
            {
               // errl.WriteLine("using" + spxt.PortName + Environment.NewLine);
            }
            #endregion
           
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }




    }
}
