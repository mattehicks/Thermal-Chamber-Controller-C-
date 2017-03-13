using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO.Ports;
namespace WindowsFormsApplication1
{



    static class Program
    {

        [STAThread]
        static void Main()
        {
            Form myForm = new Form();
            Application.Run(new Form1());
           
        }

    }

 

}