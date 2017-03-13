/* Environmental Chamber Controller
 * Compatible with select Test Equity Environmental Chambers
 * DEVELOPED BY MATT HICKS
 * FEBRUARY 2012
 * CONTACT MATTEHICKS@YAHOO.COM
 * */

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Threading;
using System.Media;

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        enum letters { A = 49, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z };
        public int MAXPLOTNUM = 36000;                        //MAX NUMBER OF DATA POINTS ON CHART
        public Graphics myG;
        public PointF g_pt;
        public ArrayList g_data = new ArrayList(36000);       //temperature graph points
        public ArrayList g_markers = new ArrayList(100);     //mark specific events
        public int g_int = 0;                               //the current number of plotted points

        Pen myPen = new Pen(Brushes.Red, 6);
        public Color universal_color;                       //to change color of a point
        public string displaytime = "step";                 //the time period to display

        public int current_x;
        public int plot_x = 1;                              //current time coordinate
        public float current_y;
        public float highest_point = 0;
        public float lowest_point = 0;

        public int scale_x = 1;
        public float scale_y = 1;
        public int marks = 0;
        public int App_interval = 5000;

        public Profile currealprofile;
        public string selectedfile;
        int curstep;
        public string oven_temp;

        DateTime start;
        DateTime full_profile;
        DateTime remaining;

        bool editmode = false;
        bool holdmode = false;
        bool monitor_temp = false;
        bool time_calculated = false;

        bool last_was_soak = false;
        bool this_is_soak = false;

        public System.Timers.Timer timer2;
        public TreeNode activeNode;
        public TreeNode Profiles;

        ContextMenuStrip panel1_menu;
        ContextMenuStrip step_change;

        //plotting variables
        public int const_msb, const_lsb;
        public string left, right;
        public sbyte signed;
        public bool neg; 

        public byte slaveid = 1;
        public byte function;
        public short MyData; 
        public SerialPort sp;
        public string error_message;

                //Message is 1 addr + 1 fcn + 2 register + 2 data + 2 CRC
                byte[] message = new byte[8];

                //Function response is fixed at 8 bytes
                public byte[] response = new byte[8];

         TextWriter errl; //error log

        public Form1()
        {
            InitializeComponent();
            create_tree();

            sp = new SerialPort("COM5", 9600);
            sp.ReadTimeout = 100;
            sp.WriteTimeout = 200;

            myG = panel1.CreateGraphics();
            myPen.DashCap = System.Drawing.Drawing2D.DashCap.Round;
            myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            universal_color = Color.Red;
            //Edit mode is disabled by default
            setpointbox.Enabled = false;
            stepbox.Enabled = false;
            ratebox.Enabled = false;
            timebox.Enabled = false;
            ratelabel .Text = "Rate" + Environment.NewLine + "  °C/min";
            panel1_menu = new ContextMenuStrip();
            ToolStripMenuItem clear = new ToolStripMenuItem("clear");
            clear.Click += new EventHandler(clear_click);
            panel1_menu.Items.AddRange(new ToolStripMenuItem[] { clear});

            /*
            timer2 = new System.Timers.Timer();
            timer2.Interval = App_interval;
            timer2.Enabled = true;
            timer2.Elapsed += new System.Timers.ElapsedEventHandler(timer2_Elapsed);
            timer2.SynchronizingObject =
            */

            step_change = new ContextMenuStrip();
            ToolStripMenuItem soak = new ToolStripMenuItem("soak");
            ToolStripMenuItem ramprate = new ToolStripMenuItem("ramprate");
            ToolStripMenuItem ramptime = new ToolStripMenuItem("ramptime");
            ToolStripMenuItem jump = new ToolStripMenuItem("jump");
            soak.Click += new EventHandler(soak_click);
            ramprate.Click += new EventHandler(ramprate_click);
            ramptime.Click += new EventHandler(ramptime_click);
            jump.Click += new EventHandler(jump_click);
           step_change.Items.AddRange(new ToolStripMenuItem[] { soak , ramprate ,ramptime,jump});

        }
        /*
        void timer2_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //IF GRAPH ON THE X AXIS done
            if ((plot_x > MAXPLOTNUM) || (g_int > MAXPLOTNUM))
            {
                if (panel3.Enabled == true)
                {
                    plot_x = 1; g_data.Clear(); g_markers.Clear(); g_int = 0;
                    g_data.Add(new DataPoint(Color.Orange, new PointF(1, g_pt.Y),System.DateTime.Now));
                    myG.Clear(Color.White);
                    panel1.BackgroundImage = WindowsFormsApplication1.Properties.Resources.graph_paper;
                }
            }

            if (monitor_temp == true)
            {
                #region set colors
                try
                {
                    WriteFunction(01, 03, (ushort)200, 01);
                    int onoff = ((int)response[4]);
                    switch (onoff)
                    {
                        case 0: statusLabel.Text = "Chamber is idle";
                            universal_color = Color.LightPink;
                            break;
                        case 2: statusLabel.Text = "Monitoring";
                            universal_color = Color.Red;
                            myPen.DashCap = System.Drawing.Drawing2D.DashCap.Round;
                            break;
                        case 3: statusLabel.Text = "Holding Profile";
                            universal_color = Color.Pink;
                            break;
                    }
                }
                catch (Exception) 
         *          { 
         *          //no response from oven?  
         *          }
                #endregion
  

                #region CHECK FOR ALARMS

                WriteFunction(01, 03, 102, 01);     //get alarm1 status 
                int Alarm_1 = response[4];
                WriteFunction(01, 03, 106, 01);     //get alarm1 status 
                int Alarm_2 = response[4];

                if (holdmode == false)
                {
                    if ((Alarm_1 == 1) || (Alarm_2 == 1))
                    {
                        holdbutton_Click(sender, null);
                        myPen.Color = Color.OrangeRed;

                        Form textwindow = new Form();
                        System.Windows.Forms.TextBox output = new System.Windows.Forms.TextBox();

                        output.Multiline = true;
                        output.Width = this.Width;
                        output.Height = this.Height;
                        textwindow.Controls.Add(output);
                        System.Drawing.Font BigPen = new System.Drawing.Font("Times New Roman", 24.0f);
                        System.Drawing.Font Smallpen = new System.Drawing.Font("Times New Roman", 18.0f);
                        // Set Font property and then add a new Label.
                        output.Font = BigPen;
                        output.Text = "Temperature Limit Reached" + Environment.NewLine + Environment.NewLine;
                        textwindow.Size = new Size(500, 300);
                        output.BackColor = Color.OrangeRed;
                        textwindow.Left = ClientRectangle.Width /2;
                        textwindow.Top = 250;
                        textwindow.Show();
                        output.Font = Smallpen;
                        output.AppendText(Environment.NewLine + "Please adjust the limits, or change the setpoint");
                        output.AppendText(Environment.NewLine + "The current profile has been paused.");

                        output.AppendText(Environment.NewLine + Environment.NewLine + 
                            "To set alarm limits goto:" + Environment.NewLine + "Tools>Operations>Temperature Limits");
                        
                    }
                    else
                    { myPen.Color = Color.Red; }
                }
                #endregion

                #region GET CURRENT STEP FROM CHAMBER
                //GET CURRENT STEP FROM CHAMBER
                    try
                    {
                        WriteFunction(01, 03, (ushort)4101, 01); //get step number
                        label1.Visible = true;
                        label1.Text = "Step: " + (response[4]).ToString();
                    }
                    catch (Exception) { };
                    label2.Visible = true;

                    try
                    {

                        WriteFunction(01, 03, (ushort)4102, 01); //get step type
                        int special = ((int)response[4]);

                        if (holdmode == false)
                        {
                            switch (special)
                            {
                                case 03: statusLabel.Text = "Soaking";
                                    break;
                                case 02: statusLabel.Text = "Ramping";
                                    break;
                            }
                        }
                    }
                    catch (Exception) { }
                #endregion

                #region GET TIME REMAINING
                    try
                    {
                        //GET TIME REMAINING FROM CHAMBER
                        if (statusLabel.Text == "Chamber is idle")
                        { label2.Visible = false; label1.Visible = false; }
                        else
                        {
                            if (displaytime == "step") //the label is clickable; switches how much time to display
                            {
                                WriteFunction(01, 03, (ushort)4120, 01);
                                int special_time = ((int)response[4] + 1);
                                label2.Text = special_time.ToString() + "min. Remaining";
                            }
                            if (displaytime == "full")
                            {
                                WriteFunction(01, 03, (ushort)4120, 01);
                                int special_time = ((int)response[4] + 1);
                                label2.Text = special_time.ToString() + "min. Remaining";
                            }
                        }
                    }
                    catch (Exception) { }
                    #endregion

                #region GET CURRENT TEMPERATURE AND CONVERT TO INTS
                    //read temp. register
                    WriteFunction(01, 03, (ushort)100, (short)01);
                    try
                    {
                        #region Positive Temps.
                        //POSITIVE TEMPERATURES
                        if (!(response[3] == 255))
                        {
                            neg = false;
                            const_msb = short.Parse(response[3].ToString()) * 256;
                            const_lsb = const_msb + short.Parse(response[4].ToString());
                            oven_temp = const_lsb.ToString();
                            int clsb = const_lsb / 10;

                            //IF TEMPERATURE IS 3 DIGITS OR MORE
                            right = oven_temp.ToString().Substring(oven_temp.Length - 1, 1);
                            if (const_lsb > 99) { tempLabel.Text = clsb + "." + right; }

                           //IF TEMPERATURE IS 2 DIGITS
                            else { tempLabel.Text = clsb + "." + right; } //if it aint broke string.format"{00:0.0}"

                            current_y = float.Parse(tempLabel.Text);//clsb;
                        }
                        #endregion

                        #region Negative Temperatures
                        //NEGATIVE TEMPERATURES
                        else
                        {
                            //CONVERTING FROM TWOS COMPLEMENT
                            //CONVERSION BYTE IS LESS THAN 128, IMPLICIT CONVERSION CAST
                            byte signedT = (byte)response[4];
                            if (signedT < -128)
                            {
                                string decade = (signedT / 10).ToString();
                                string s_right = signedT.ToString();

                                right = signedT.ToString().Substring(s_right.Length - 1, 1);
                                if (signedT < 99)
                                {
                                    tempLabel.Text = decade + "." + right;
                                }
                                else
                                {
                                    tempLabel.Text = signedT + "." + right;
                                }
                                neg = true;
                                current_y = signedT / 10;
                            }

                            //CONVERSION IS EXPLICIT, VALUES OVER 128 WILL OVERWRITE THE MSB WHICH HOLDS THE NEGATIVE BIT
                            if (signedT >= -128)
                            {
                                //the exact function as above, except values are subtracted from a reversed empty byte (255), and added to 12.8
                                //const_msb = short.Parse(response[3].ToString()) * 256; THE VALUE IS NOT EXPECTED TO PASS 128 DEGREES
                                short big_neg_num = (short)(128 + (128 - signedT)); //a formatted representation of 12.8 degrees plus actual twos compliment bits received
                                int big_divided_num = big_neg_num / 10;             // the whole part of the degrees
                                //int big_decimal_num = 128 - signedT;                // the decimal part of the degrees

                                string big_right = big_neg_num.ToString().Substring(big_neg_num.ToString().Length - 1, 1);

                                tempLabel.Text = "-" + big_divided_num + "." + big_right;
                                neg = true;
                                current_y = ~(big_neg_num / 10);
                            }
                        }

                        #endregion

                        panel2.Refresh();
                        tempLabel.Refresh();
                    }
                    catch { };
                    #endregion

                #region POST TO GUI
                    try
                    {
                        if (g_data.Count == 0)
                        {
                            plot_x = 1;
                            current_x = plot_x;
                            g_int = 0;
                            myG.Clear(Color.White);
                            panel1.BackgroundImage = WindowsFormsApplication1.Properties.Resources.graph_paper;
                            g_pt = new PointF(0, 0);
                            g_data.Add(new DataPoint(universal_color, g_pt,DateTime.Now));//starting point on graph
                        }

                        else
                        {
                            //CALCULATE THE CURRENT POINT
                            current_x = plot_x;
                            g_int += 1;
                            g_pt = new PointF(current_x, current_y);
                            g_data.Add(new DataPoint(universal_color, g_pt,DateTime.Now));

                            //DRAW THE TEMPERATURE POINTS
                            for (int i = 1; i < g_data.Count; i++)
                            {
                                    DataPoint curdatapoint = (DataPoint)g_data[i];
                                    myPen.Color = curdatapoint.calorie;

                                    DataPoint temppoint = (DataPoint)g_data[i-1];
                                    DataPoint olddot = new DataPoint(temppoint.calorie, temppoint.pit,DateTime.Now);

                                    float midpoint = panel1.Height / 2;
                                    float cury = curdatapoint.pit.Y / 60;
                                    float mappoint = midpoint - (cury * midpoint);
                                    float oldy = olddot.pit.Y / 60;
                                    float oldmappoint = midpoint - (oldy * midpoint);
                                    PointF p1 = new PointF(i / scale_x - 1, oldmappoint * scale_y);
                                    PointF p2 = new PointF(i / scale_x, mappoint * scale_y);

                                    myG.DrawLine(myPen, p1, p2);
                            }

                            // DRAW THE START/STOP MARKERS
                            foreach (DataPoint marker in g_markers)
                            {
                                float midpoint = panel1.Height / 2;
                                float cury = marker.pit.Y / 60;
                                float mappoint = midpoint - (cury * midpoint);

                                myG.DrawEllipse(new Pen(Brushes.ForestGreen, 5), new System.Drawing.Rectangle(new System.Drawing.Point((int)marker.pit.X / scale_x, (int)(mappoint *scale_y)), new Size(7, 7)));
                            }
                          }
                        }
                    catch (Exception) {  /* statusLabel.Text = "PlotData fail"; }
                #endregion 

                            default_loop();
                    plot_x++;
                }
            else { timer2.Stop(); }

        }
        */
        private void clear_click(object sender, EventArgs e)
        {
            if (panel3.Enabled == true)
            {
                //right mouse click handler
                g_data.Clear();
                g_markers.Clear();
                tempLabel.Text = " --";
                g_int = 0;
                plot_x = 1;
                myG.Clear(Color.White);
                panel1.BackgroundImage = WindowsFormsApplication1.Properties.Resources.graph_paper;
            }
            else { }

        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (monitor_temp == false) //Setup the GUI and turn on
            {
                #region open serial port
                if (!(sp.IsOpen))
                {
                    try
                    {
                        sp.Open();
                    }
                    catch (Exception)
                    {
                        try
                        {
                            sp = new SerialPort("COM4", 9600);
                            sp.Open();
                        }
                        catch (Exception)
                        {
                            try
                            {
                                sp = new SerialPort("COM6", 9600);
                                sp.Open();
                            }
                            catch (Exception)
                            {
                                checkBox1.Checked = false;
                                MessageBox.Show("COM4,5,6 not ready", "MsgBox",
                                MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                            }
                        }
                    }

                }
                else { /* success */};
                #endregion

                //everything is OK.
                clear_click(this, null);
                monitor_temp = true;
                default_loop();
                statusLabel.Text = "Monitoring";
                commStatusLabel.Text = "SerPt Open";
            }
            else
            {
                //Turn off the monitor
                sp.Close();
                statusLabel.Text = "Not Monitoring"; 
                commStatusLabel.Text = "SerPt Closed";
                label1.Visible = false;
                label2.Visible = false;
                monitor_temp = false;
            }
        }

        public void default_loop()
        {
            timer1.Interval = App_interval;
            timer1.Start();
        }

        public void soak_click(object sender, EventArgs e)
        {
            StepData tempstep;

            if (!(activeNode == null))
            {
                if (!(treeView1.SelectedNode == null))
                {
                    tempstep = (StepData)currealprofile.data[curstep-1];
                    tempstep.type = "soak";
                    create_tree();
                    update_controls(curstep);
                }
            }
        }

        public void ramprate_click(object sender, EventArgs e)
        {
            StepData tempstep;

            if (!(activeNode == null))
            {
                if (!(treeView1.SelectedNode == null))
                {
                    tempstep = (StepData)currealprofile.data[curstep - 1];
                    tempstep.type = "ramprate";
                    create_tree();
                    update_controls(curstep);
                }
            }
        }

        public void ramptime_click(object sender, EventArgs e)
        {
            StepData tempstep;

            if (!(activeNode == null))
            {
                if (!(treeView1.SelectedNode == null))
                {
                    tempstep = (StepData)currealprofile.data[curstep - 1];
                    tempstep.type = "ramptime";
                    create_tree();
                    update_controls(curstep);
                }
            }
        }

        public void jump_click(object sender, EventArgs e)
        {
            if (!(activeNode == null))
            {
                if (!(treeView1.SelectedNode == null))
                {
                    jumpStepToolStripMenuItem_Click(sender, e);
                    create_tree();
                    update_controls(curstep);
                }
            }
        }

        public void JumpInsert(object sender, EventArgs e, int file, int step, int repeat)
        {
            if (!(currealprofile == null))
            {
                //jumpstep params:  File , Step , Repeat. //change the variable 5 to the number in 'file'
                StepData newstep = new StepData("jump", (short)5, (byte)step, (byte)repeat);
                currealprofile.data.Insert(curstep-1, newstep);
                currealprofile.data.RemoveAt(curstep);
                create_tree();
                update_controls(curstep);
            }
        }

        void create_tree()
        {
            treeView1.Nodes.Clear();
            TreeNode p1n1 = new TreeNode("Step1:");
            TreeNode p1n2 = new TreeNode("Step2:");
            TreeNode p1n3 = new TreeNode("Step3:");
            TreeNode p1n4 = new TreeNode("Step4:");
            TreeNode[] atp_step_array = new TreeNode[] { p1n1, p1n2, p1n3, p1n4 };
            TreeNode tn1 = new TreeNode("ATP FULL", atp_step_array);
            //treeView1.Nodes.Add(tn1);

            TreeNode p2n1 = new TreeNode("Step1:");
            TreeNode p2n2 = new TreeNode("Step2:");
            TreeNode p2n3 = new TreeNode("Step3:");
            TreeNode p2n4 = new TreeNode("Step4:");
            TreeNode[] Parray = new TreeNode[] { p2n1, p2n2, p2n3, p2n4 };
            TreeNode tn2 = new TreeNode("One Cycle", Parray);
            //treeView1.Nodes.Add(tn2);

            if (!(currealprofile == null))
            {
                TreeNode Parent = new TreeNode(currealprofile.name.ToString());
                Parent.Name = currealprofile.name;
                //Parent Node\
                activeNode = Parent;
                int x = 1;
                foreach (StepData xint in currealprofile.data)
                {
                    TreeNode child = new TreeNode();
                    if (xint.type == "Soak" || xint.type == "soak")
                    {
                        child.Text = "Step" + x + ": " + xint.type + " " + " -- "
                           + int.Parse(xint.time.ToString()) + ":00 ";
                    }
                    else if (xint.type == "jump" || xint.type == "Jump")
                    {
                        child.Text = "Step" + x + ": " + xint.type + "    File:" + xint.temp + " " +
                           "Step:" + int.Parse(xint.time.ToString()) + " x" + xint.endvalue;
                    }
                    else
                    {
                        child.Text = "Step" + x + ": " + xint.type + " " + int.Parse(xint.temp.ToString()) + "°C "
                           + int.Parse(xint.time.ToString()) + ":00 ";
                    }
                    child.Name = "Step" + x;
                    Parent.Nodes.Add(child);
                    x++;
                }
                TreeNode[] LocalProfiles = new TreeNode[] { tn1, tn2, Parent };
                TreeNode Profiles = new TreeNode("Profiles", LocalProfiles);
                treeView1.Nodes.Add(Profiles);

                activeNode = treeView1.TopNode;

                activeNode.Expand();

                activeNode = activeNode.LastNode;

                for (int i = 0; i < treeView1.Nodes.Count; i++)
                {
                    if (!(activeNode.Name.Contains(currealprofile.name)))
                    {
                        activeNode = activeNode.NextNode;
                    }

                    else
                    {
                        statusLabel.Text = activeNode.Name.ToString();
                        activeNode.Expand();
                    }
                }
            }

            else
            {
                TreeNode[] LocalProfiles = new TreeNode[] { tn1, tn2 };
                Profiles = new TreeNode("Profiles", LocalProfiles);
                treeView1.Nodes.Add(Profiles);
            }
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //add functionality to update currprofile to the node thats selected.

            activeNode = e.Node;
            if (!(activeNode == null))
            {
                //statusLabel.Text = activeNode.Index.ToString();
                if (!(treeView1.SelectedNode == null))
                {
                    statusLabel.Text = activeNode.ToString();
                    if (e.Button == System.Windows.Forms.MouseButtons.Right)
                    {
                        step_change.Show(treeView1, new System.Drawing.Point(e.X, e.Y));
                    }

                }
                curstep = e.Node.Index + 1;                  // indices start at ZERO!!!!
                stepbox.Text = curstep.ToString();
                if (!(currealprofile == null))
                {
                    update_controls(curstep);
                }
            }
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            editmode = true;
            update_controls(activeNode.Index);
        }

        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                panel1_menu.Show(panel1, new System.Drawing.Point(e.X, e.Y));
            }
            else
            {
                if (monitor_temp == false)
                {
                    int xm = e.X;
                    int ym = e.Y;
                    g_data.Insert(g_int, new DataPoint(Color.Black, new PointF(xm, ym),System.DateTime.Now));

                    if (g_int == 0) { /* in this case we have zero data, skip rendering  */}
                    else
                    {
                        int g_old_int = g_int - 1;
                        DataPoint g_old = (DataPoint)g_data[g_old_int];
                        g_pt = new PointF(xm, ym);
                        g_data.Add(new DataPoint(universal_color, g_pt,System.DateTime.Now));

                        myG.DrawLine(myPen, g_old.pit, g_pt);
                        statusLabel.Text = (g_pt.ToString());
                    }
                    g_int += 1;
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
            displaytime = displaytime == "step" ? "full" : "step";
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            editmode = false;
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.InitialDirectory = "C:\\Desktop";
            fdlg.Title = "Profile Load";
            fdlg.Filter = "Text files (*.txt)|*.txt";

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                selectedfile = fdlg.FileName;
                TextReader textan = new StreamReader(selectedfile);
                string localname = Path.GetFileName(selectedfile);
                string blurb = textan.ReadLine();
                int localsteps = int.Parse(textan.ReadLine());
                currealprofile = new Profile(localname);
                currealprofile.steps = localsteps;
                currealprofile.name = localname;

                string tempstring;
                string[] tempdata = new string[localsteps];
                StepData tempstepdata;

                try{
                    for (int i = 0; i < localsteps; i++)
                    {
                        tempstring = textan.ReadLine();
                        tempstring = tempstring.Trim();
                        tempdata = tempstring.Split(',', '[', ']', ' ');

                        //remove all whitespace characters, otherwise send in empty data.
                        int a = 0;
                        foreach (string item in tempdata)
                        {
                            if (item.Length >= 1)
                            {
                                tempdata[a] = item;
                                a++;
                            }
                        }
                        //the data is relocated into these 3 positions.
                        if (!(tempdata[0] == "Soak"))
                        {
                            tempstepdata = new StepData(tempdata[0], short.Parse(tempdata[1]), byte.Parse(tempdata[2]), byte.Parse(tempdata[3]));
                            currealprofile.data.Add(tempstepdata);
                        }
                        else
                        {
                            tempstepdata = new StepData(tempdata[0], byte.Parse(tempdata[2]));
                            currealprofile.data.Add(tempstepdata);
                        }
                    }
                }
                    catch(Exception){}

                // profile needs to be fully loaded by this time
                curstep = 1;
                activeNode = treeView1.TopNode;
                activeNode.Expand();
                create_tree();
                update_controls(curstep);
                textan.Close();
            }
        }

        private void newToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            StepData tempstep = new StepData("ramprate", 35, 3, 0);
            currealprofile = new Profile("NewProfile", tempstep);
       
            curstep = 1;
            editmode = true;

            create_tree();
            update_controls(1);
        }

        private void exitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            sp.Close();
            this.Close();
        }

        private void savebutton_Click(object sender, EventArgs e)
        {
            // add functionality to capture profile name as read from TreeNode.name
            try
            {
                string myname = currealprofile.name;
                string[] name = myname.Split('.');
                TextWriter tw = new StreamWriter(name[0] + ".txt");

                tw.WriteLine(currealprofile.name);
                tw.WriteLine(currealprofile.steps);

                foreach (StepData itemd in currealprofile.data)
                {
                    tw.WriteLine("[" + itemd.type + "," + itemd.temp.ToString() + "," + itemd.time.ToString() + "," + itemd.endvalue.ToString() + "]");
                }
                tw.Close();
                statusLabel.Text = "saved.";
            }
            catch (Exception) { statusLabel.Text = "file error"; }
        }
        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                // add functionality to capture profile name as read from TreeNode.name
                string myname = currealprofile.name;
                string[] name = myname.Split('.');
                TextWriter tw = new StreamWriter(name[0] + ".txt");

                tw.WriteLine(currealprofile.name);
                tw.WriteLine(currealprofile.steps);

                foreach (StepData itemd in currealprofile.data)
                {
                    tw.WriteLine("[" + itemd.type + "," + itemd.temp.ToString() + "," + itemd.time.ToString() + "," + itemd.endvalue.ToString() + "]");
                }
                tw.Close();
                statusLabel.Text = "saved.";
            }
            catch (Exception) { statusLabel.Text = "file error"; }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filename = currealprofile.ToString();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "C:\\Documents and Settings\\";
            saveFileDialog1.Title = "Profile Save";
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt";

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (var writer = new StreamWriter(saveFileDialog1.FileName))
                {
                    writer.WriteLine(currealprofile.name);
                    writer.WriteLine(currealprofile.steps);

                    foreach (StepData itemd in currealprofile.data)
                    {
                        writer.WriteLine("[ " + itemd.type + ", " + itemd.temp + ", " + itemd.time + "," + itemd.endvalue + "]");
                    }
                    writer.Close();
                }
                int f_int = saveFileDialog1.FileName.LastIndexOf("\\") + 1;
                currealprofile.name = saveFileDialog1.FileName.Substring(f_int);
                create_tree();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            serialPort1.Close();
            this.Close();
        }

        private void Newbutton_Click(object sender, EventArgs e)
        {
            StepData tempstep = new StepData("ramprate", 35, 3, 0);
            currealprofile = new Profile("NewProfile", tempstep);

            curstep = 1;
            editmode = true;

            create_tree();
            update_controls(1);
        }

        private void insertbutton_Click(object sender, EventArgs e)
        {
            if (!(currealprofile == null))
            {
                StepData newstep = new StepData("ramprate", 35, 3, 0);
                currealprofile.data.Insert(curstep, newstep);
                currealprofile.steps += 1;
                create_tree();
                update_controls(curstep);
                //curstep += 1;
            }
        }

        private void deletebutton_Click(object sender, EventArgs e)
        {
            if (!(currealprofile == null))
            {
                try
                {
                    if (!(currealprofile.steps < 1)) //profile.steps.1 must match the profile.data[0]
                    {   //error case: current step is 1, and there is no data in 0.
                        //profile.data[0] exists whether or not there is data.
                        if (curstep == 0) { curstep = currealprofile.steps; }
                        currealprofile.data.RemoveAt(curstep - 1); //transform from GUI space to Array(data) Space
                        currealprofile.steps -= 1;
                        statusLabel.Text = "removed:" + curstep;
                        create_tree();
                        update_controls(curstep);
                        curstep -= 1;
                        if (curstep > currealprofile.steps) { curstep = currealprofile.steps + 1; }
                    }
                }
                catch (Exception) { curstep = currealprofile.steps; }
            }
        }

        private void UpButton_Click(object sender, EventArgs e)
        {
            if (!(currealprofile == null))
            {
                if (curstep < currealprofile.steps)
                {
                    curstep += 1;
                    stepbox.Text = (curstep).ToString();
                    update_controls(curstep);
                }
                else
                {
                    //do nothing
                }
            }
            else
            {
                MessageBox.Show("Select a profile to start", "MsgBox",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            }
        }

        private void DownButton_Click(object sender, EventArgs e)
        {
            if (!(currealprofile == null))
            {
                if (curstep > 1)
                {
                    curstep -= 1;
                    stepbox.Text = (curstep ).ToString();
                    update_controls(curstep);
                }
                else
                {
                    // do nothing
                }
            }
            else
            {
                MessageBox.Show("Select a profile to start", "MsgBox",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            }
        }

        private void Zoombutton2_Click(object sender, EventArgs e)
        {
            if (!(scale_x > 10)) { scale_x++; label_scale.Visible = true; label_scale.Text = "X" + scale_x.ToString(); }
        }

        private void Zoombutton1_Click(object sender, EventArgs e)
        {
            if (!(scale_x <= 1)) { scale_x--; label_scale.Visible = true; label_scale.Text = "X" + scale_x.ToString(); }
        }

        private void stepbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(!(currealprofile == null))
            {
            if (e.KeyChar == (char)Keys.Return)
            {
                        curstep  = int.Parse(stepbox.Text);
                        update_controls(curstep);
              }
                    else
                    {
                        //do nothing
                    }
            }
                else
                {
                    MessageBox.Show("Select a profile to start", "MsgBox",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }
            }

        private void setpointbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Will change the temperature set point in the profile---- WITHOUT SAVING
            if (e.KeyChar == (char)Keys.Return)
            {
                if (!(currealprofile == null))
                {//begin data trap for undo method - not
                    if (!(curstep >= 1)) { curstep = 1; goto end; }

                    if (activeNode == null)
                    {
                        MessageBox.Show("Select a profile to start", "MsgBox",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        goto end;
                    }

                    #region store TreeNode in tempProfile
                    if (!(activeNode.Parent.Name == currealprofile.name))
                    {
                        #region no code
                        statusLabel.Text = "Pick Desired Profile";
                        #endregion
                    }
                    #endregion

                    #region editmode
                    if (editmode == true)
                    {
                        if (setpointbox.Text == "")
                        {
                            goto end;
                        }
                        //else
                        stepbox.Text = curstep.ToString(); //we now have a step number
                        short val_part = short.Parse(setpointbox.Text.ToString());
                        //GET and SET
                        StepData localstepdata;
                        localstepdata = (StepData)currealprofile.data[curstep - 1];
                        localstepdata.temp = val_part;
                        currealprofile.data[curstep - 1] = localstepdata;

                        statusLabel.Text = "changed: " + curstep;
                        create_tree();
                        update_controls(curstep);

                        //add type change handler
                    }
                    #endregion
                }

                else
                {
                    MessageBox.Show("Load a profile to start", "MsgBox",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }

            end:
                { //ugggghhhhh}
                }
            }
        }

        private void timebox_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Will change the temperature set point in the profile---- WITHOUT SAVING
            if (e.KeyChar == (char)Keys.Return)
            {
                if (!(currealprofile == null))
                {//begin data trap for undo method
                    if (!(curstep >= 1)) { curstep = 1; goto end; }

                    if (activeNode == null)
                    {
                        MessageBox.Show("Select a profile to start", "MsgBox",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        goto end;
                    }

                    #region store TreeNode in tempProfile
                    if (!(activeNode.Parent.Name == currealprofile.name))
                    {
                        #region no code
                        statusLabel.Text = "Pick Desired Profile";
                        #endregion
                    }
                    #endregion

                    #region editmode
                    if (editmode == true)
                    {
                        if (timebox.Text == "")
                        {
                            goto end;
                        }
                        //else
                        //GET and SET
                        StepData localstepdata;
                        localstepdata = (StepData)currealprofile.data[curstep - 1];

                        if ((localstepdata.type == "soak") || (localstepdata.type == "Soak"))
                        {
                            byte time_part = byte.Parse(timebox.Text.ToString());
                            localstepdata.time = time_part;
                        }
                        else
                        {
                            byte time_part = byte.Parse(timebox.Text.ToString());//this is okay for ramp rate values
                            localstepdata.time = time_part;
                            short temp_part = short.Parse(setpointbox.Text.ToString());
                            localstepdata.temp = temp_part;
                        }
  
                        currealprofile.data[curstep - 1] = localstepdata;
                        //done

                        create_tree();
                        update_controls(curstep);
                        statusLabel.Text = "changed: " + curstep;
                        //add type change handler
                    }
                    #endregion
                }

                else
                {
                    MessageBox.Show("Load a profile to start", "MsgBox",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }

            end:
                { //ugggghhhhh}
                }
            }
        }

        private void ratebox_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Will change the temperature set point in the profile---- WITHOUT SAVING
            if (e.KeyChar == (char)Keys.Return)
            {
                if (!(currealprofile == null))
                {//begin data trap for undo method
                    if (!(curstep >= 1)) { curstep = 1; goto end; }

                    if (activeNode == null)
                    {
                        MessageBox.Show("Select a profile to start", "MsgBox",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        goto end;
                    }

                    #region store TreeNode in tempProfile
                    if (!(activeNode.Parent.Name == currealprofile.name))
                    {
                        #region no code
                        statusLabel.Text = "Pick Desired Profile";
                        #endregion
                    }
                    #endregion

                    #region editmode
                    if (editmode == true)
                    {
                        if (ratebox.Text == "")
                        {
                            goto end;
                        }
                        //else
                        //GET and SET
                        StepData localstepdata;
                        localstepdata = (StepData)currealprofile.data[curstep - 1];

                        if ((localstepdata.type == "ramprate") || (localstepdata.type == "Ramprate"))
                        {
                            byte time_part = byte.Parse(ratebox.Text.ToString());
                            localstepdata.time = time_part;
                        }
                        else
                        {
                            //this should not occur, rampbox is disabled.
                            /*
                            byte time_part = byte.Parse(timebox.Text.ToString());//this is okay for ramp rate values
                            localstepdata.time = time_part;
                            short temp_part = short.Parse(setpointbox.Text.ToString());
                            localstepdata.temp = temp_part;
                             */
                        }

                        currealprofile.data[curstep - 1] = localstepdata;
                        //done

                        create_tree();
                        update_controls(curstep);
                        statusLabel.Text = "changed: " + curstep;
                        //add type change handler
                    }
                    #endregion
                }

                else
                {
                    MessageBox.Show("Load a profile to start", "MsgBox",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }

            end:
                { //ugggghhhhh}
                }
            }
        }

        private void TerminateButton_Click(object sender, EventArgs e)
        {
            if (sp.IsOpen == false) { try { EstablishConnection(); } catch (Exception) { }; }

            WriteFunction(01, 06, 1217, (short)01); holdmode = true;// terminate
            statusLabel.Text = "Terminating";
        }

        private void holdbutton_Click(object sender, EventArgs e)
        {
            if (holdmode == false)
            {
                WriteFunction(01, 06, 1210, (short)01); holdmode = true;// pause
                universal_color = Color.Pink;
                statusLabel.Text = "Holding";
                holdbutton.Text = "resume";
                holdmode = true;
            }
                else{
                    WriteFunction(01, 06, 1209, (short)01); holdmode = true; // resume
                    universal_color = Color.Red;
                    statusLabel.Text = "Resuming";
                    holdbutton.Text = "Hold";
                    myPen.Color = Color.Red;
                    holdmode = false;
                }
        }

        private void StartButton_Click_1(object sender, EventArgs e)
        {
            if (!(currealprofile == null))
            {
                editmode = false;
                curstep = 1;

                errl = new StreamWriter("ThermalLOG.txt");
                errl.WriteLine(System.DateTime.Now + Environment.NewLine);

                try
                {
                    //start communication with F4 Control
                    #region start upload
                    if (!(currealprofile.steps == 0) || !(currealprofile.data.Count == 0))
                    {
                        /***************************************/
                        // MAKE ROOM IN THE CONTROLLER MEMORY, AKA DELETE FILE 5
                        error_message = "delete file 1";
                        WriteFunction(01, 06, 4000, 5); //FILE
                        error_message = "delete file 2";
                        WriteFunction(01, 06, 4001, 1); //STEP
                        error_message = "delete file 3";
                        WriteFunction(01, 06, 4002, 3); //DELETE
                        error_message = "delete file 4";
                        //WriteFunction(01, 06, 0025, 00); //SAVE causes an error because it does not have a response

                        //  WRITE A FILE TO THE MEMORY
                        error_message = "file pointer";
                        WriteFunction(01, 06, 4000, 5);

                        error_message = "create profile";
                        WriteFunction(01, 06, 4002, 1); //command: create profile

                        error_message = "naming x4"; //file 5 start, 3500+(10 * filenumber) 
                        WriteFunction(01, 06, 3540, 80); // rename????? P 
                        WriteFunction(01, 06, 3541, 65); // rename????? A
                        WriteFunction(01, 06, 3542, 83); // rename????? S
                        WriteFunction(01, 06, 3543, 83); // rename????? S
                        WriteFunction(01, 06, 3544, 80); // rename????? P
                        WriteFunction(01, 06, 3545, 82); // rename????? R
                        WriteFunction(01, 06, 3546, 79); // rename????? O
                        WriteFunction(01, 06, 3547, 70); // rename????? F
                        WriteFunction(01, 06, 3548, 73); // rename????? I   
                        WriteFunction(01, 06, 3549, 76); // rename????? L

                        int i;
                        StepData isegment;
                        for (i = 1; i <= currealprofile.steps; i++)
                        {
                            isegment = (StepData)currealprofile.data[i - 1];

                            if ((isegment.type == "soak") || (isegment.type == "Soak"))
                            {
                                error_message = "soakstep" + i;
                                WriteFunction(01, 06, 4001, (short)i); //set step#
                                WriteFunction(01, 06, 4002, (short)2); // insert step
                                WriteFunction(01, 06, 4012, (short)0); // clear wait flag
                                WriteFunction(01, 06, 4003, (short)03); // type soak
                                WriteFunction(01, 06, 4010, (short)(isegment.time));
                                error_message = "save soak" + i;
                                WriteFunction(01, 06, 0025, (short)00);  //save step
                            }

                            if (isegment.type == "ramprate")
                            {
                                error_message = "ramprate" + i;
                                WriteFunction(01, 06, 4001, (short)i); //set step#
                                WriteFunction(01, 06, 4002, (short)2); // insert step
                                WriteFunction(01, 06, 4012, (short)0); // clear wait flag

                                WriteFunction(01, 06, 4003, (short)02); // type ramp rate
                                WriteFunction(01, 06, 4043, (short)(isegment.time * 10)); //  degrees per minute
                                WriteFunction(01, 06, 4044, (short)(isegment.temp * 10)); // setpoint
                                error_message = "save ramprate" + i;
                                WriteFunction(01, 06, 0025, (short)0);  //save step
                            }

                            if (isegment.type == "ramptime")
                            {
                                error_message = "ramptime" + i;
                                WriteFunction(01, 06, 4001, (short)i); //set step#
                                WriteFunction(01, 06, 4002, (short)2); // insert step
                                WriteFunction(01, 06, 4012, (short)0); // clear wait flag

                                WriteFunction(01, 06, 4003, (short)01); // type ramp-time
                                WriteFunction(01, 06, 4010, (short)(isegment.time)); //  ramp time
                                WriteFunction(01, 06, 4044, (short)(isegment.temp * 10)); // setpoint
                                error_message = "save ramptime" + i;
                                WriteFunction(01, 06, 0025, 0);  //save step
                            }

                            if (isegment.type == "jump")
                            {
                                error_message = "jump" + i;
                                WriteFunction(01, 06, 4001, (short)i); //set step#
                                WriteFunction(01, 06, 4002, (short)2); // insert step
                                WriteFunction(01, 06, 4012, (short)0); // clear wait flag

                                WriteFunction(01, 06, 4003, (short)04); // type jump
                                WriteFunction(01, 06, 4050, (short)isegment.temp); //  jump to file
                                WriteFunction(01, 06, 4052, (short)(isegment.endvalue)); // repeat
                                error_message = "save jump" + i;
                                WriteFunction(01, 06, 0025, 0);  //save step
                            }

                            error_message = "clear flag";
                            WriteFunction(01, 06, 4002, 0); // clear insert flag
                            //WriteFunction(01, 06, 25, 00);  //save step
                            /***********************************************/
                        }

                        //set end step
                        error_message = "end step x3";
                        WriteFunction(01, 06, 4001, (short)(i + 1));
                        WriteFunction(01, 06, 4002, (short)02); // insert step
                        WriteFunction(01, 06, 4060, (short)03); // end idle
                        //WriteFunction(01, 06, 4060, (short)00); // end hold

                        //start profile
                        error_message = "start profile x3";
                        WriteFunction(01, 06, 4000, (short)5); //set file position in F4 controller
                        WriteFunction(01, 06, 4001, (short)1);  //set pointer to step one
                        WriteFunction(01, 06, 4002, (short)5);  //start
                        #endregion
                        commStatusLabel.Text = "Upload Complete";
                    }
                    else
                    {
                        //no currsteps 
                        commStatusLabel.Text = "check file";
                    }
                }
                catch (TimeoutException)
                {
                    commStatusLabel.Text = "timed out";
                }
            }
            else
            {
                if (monitor_temp == true) { default_loop(); }
                else
                {
                    MessageBox.Show("Select a profile to start", "MsgBox",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }
            }

            try
            {
                errl.Close();
                statusLabel.Text = "Running: " + currealprofile.name.ToString();
                    default_loop();
                    //start button marker
                    g_markers.Insert(g_markers.Count,new DataPoint(Color.Gold, g_pt,System.DateTime.Now));
            }
            catch { statusLabel.Text = "Error"; };
        }

        public void timer1_Tick(object sender, EventArgs e)
        {
            //If the GUI is full---
            if ((plot_x > MAXPLOTNUM) || (g_int > MAXPLOTNUM))
            {
                if (panel3.Enabled == true) //If the controls have been locked
                {
                    plot_x = 1; g_data.Clear(); g_markers.Clear(); g_int = 0;
                    g_data.Add(new DataPoint(Color.Orange, new PointF(1, g_pt.Y), System.DateTime.Now));
                    myG.Clear(Color.White);
                    panel1.BackgroundImage = WindowsFormsApplication1.Properties.Resources.graph_paper;
                }
                else {
                    try
                    {
                        myG.DrawString("Please unlock or clear the GUI", new System.Drawing.Font(FontFamily.GenericSansSerif,20),Brushes.Black,new System.Drawing.Point(100,100));
                        Console.Beep(400, 100);
                    }
                    catch (Exception)
                    {
                        myG.DrawString("Please unlock or clear the GUI", new System.Drawing.Font(FontFamily.GenericSansSerif, 20), Brushes.Black, new System.Drawing.Point(100, 100));
                         SystemSounds.Beep.Play();
                    }
                }
            }

            if (monitor_temp == true)
            {
                #region set colors
                try
                {
                    WriteFunction(01, 03, (ushort)200, 01);
                    int onoff = ((int)response[4]);
                    switch (onoff)
                    {
                        case 0: statusLabel.Text = "Chamber is idle";
                            universal_color = Color.LightPink;
                            break;
                        /*case 2: statusLabel.Text = "Monitoring";
                            universal_color = Color.Red;
                            myPen.DashCap = System.Drawing.Drawing2D.DashCap.Round;
                            break;
                        */
                        case 3: statusLabel.Text = "Holding Profile";
                            universal_color = Color.Pink;
                            break;
                    }
                }
                catch (Exception) { /*no response from oven? */ }
                #endregion

                #region CHECK FOR ALARMS

                WriteFunction(01, 03, 102, 01);     //get alarm1 status 
                int Alarm_1 = response[4];
                WriteFunction(01, 03, 106, 01);     //get alarm1 status 
                int Alarm_2 = response[4];

                if (holdmode == false)
                {
                    if ((Alarm_1 == 1) || (Alarm_2 == 1))
                    {
                        holdbutton_Click(sender, e);
                        myPen.Color = Color.OrangeRed;

                        Form textwindow = new Form();
                        System.Windows.Forms.TextBox output = new System.Windows.Forms.TextBox();

                        output.Multiline = true;
                        output.Width = this.Width;
                        output.Height = this.Height;
                        textwindow.Controls.Add(output);
                        System.Drawing.Font BigPen = new System.Drawing.Font("Times New Roman", 24.0f);
                        System.Drawing.Font Smallpen = new System.Drawing.Font("Times New Roman", 18.0f);
                        // Set Font property and then add a new Label.
                        output.Font = BigPen;
                        output.Text = "Temperature Limit Reached" + Environment.NewLine + Environment.NewLine;
                        textwindow.Size = new Size(500, 300);
                        output.BackColor = Color.OrangeRed;
                        textwindow.Left = ClientRectangle.Width /2;
                        textwindow.Top = 250;
                        textwindow.Show();
                        output.Font = Smallpen;
                        output.AppendText(Environment.NewLine + "Please adjust the limits, or change the setpoint");
                        output.AppendText(Environment.NewLine + "The current profile has been paused.");

                        output.AppendText(Environment.NewLine + Environment.NewLine + 
                            "To set alarm limits goto:" + Environment.NewLine + "Tools>Operations>Temperature Limits");
                        
                    }
                    else
                    { myPen.Color = Color.Red; }
                }
                #endregion

                #region GET CURRENT STEP FROM CHAMBER
                //GET CURRENT STEP FROM CHAMBER
                    try
                    {
                        WriteFunction(01, 03, (ushort)4101, 01); //get step number
                        label1.Visible = true;
                        label1.Text = "Step: " + (response[4]).ToString();
                    }
                    catch (Exception) { };
                    label2.Visible = true;

                    try
                    {
                        WriteFunction(01, 03, (ushort)4102, 01); //get step type
                        int special = ((int)response[4]);

                        if (holdmode == false)
                        {
                            switch (special)
                            {
                                case 03: statusLabel.Text = "Soaking";
                                    break;
                                case 02: statusLabel.Text = "Ramping";
                                    break;
                            }
                        }
                    }
                    catch (Exception) { }
                #endregion

                #region GET TIME REMAINING
                        try
                        {
                            //GET TIME REMAINING FROM CHAMBER
                            if (statusLabel.Text == "Chamber is idle")
                            { label2.Visible = false; label1.Visible = false; }
                            else
                            {
                                //KILL ALL THIS STUFF, CALCULATE MANUALLY

                                if (displaytime == "step") //the label is clickable; switches how much time to display
                                {
                                    WriteFunction(01, 03, (ushort)4120, 01);
                                    int short_time = ((int)response[4] + 1);
                                    label2.Text = short_time.ToString() + "min. Remaining";
                                }
                                if (displaytime == "full")
                                {
                                   
                                  if(time_calculated == false) //has the long time been calculated
                                  { calculate_time(); }
                                  else
                                  {
                                      label2.Text = remaining.ToString() + "min. Remaining";
                                  }
                                }
                            }
                        }
                        catch (Exception) { }
                #endregion

                #region GET CURRENT TEMPERATURE AND CONVERT TO INTS
                    //read temp. register
                    WriteFunction(01, 03, (ushort)100, (short)01);
                    try
                    {
                        #region Positive Temps.
                        //POSITIVE TEMPERATURES
                        if (!(response[3] == 255))
                        {
                            neg = false;
                            const_msb = short.Parse(response[3].ToString()) * 256;
                            const_lsb = const_msb + short.Parse(response[4].ToString());
                            oven_temp = const_lsb.ToString();
                            int clsb = const_lsb / 10;

                            //IF TEMPERATURE IS 3 DIGITS OR MORE
                            right = oven_temp.ToString().Substring(oven_temp.Length - 1, 1);
                            if (const_lsb > 99) { tempLabel.Text = clsb + "." + right; }

                           //IF TEMPERATURE IS 2 DIGITS
                            else { tempLabel.Text = clsb + "." + right; } //if it aint broke string.format"{00:0.0}"

                            current_y = float.Parse(tempLabel.Text); //clsb

                            if (statusLabel.Text == "Soaking")
                            {
                                this_is_soak = true;
                                if ((this_is_soak == true) && (last_was_soak == false)) //logic; first timer tick at new soak step
                                {
                                    g_markers.Add(null);
                                }

                                float num = float.Parse(tempLabel.Text);
                                if (num > highest_point)
                                {
                                    highest_point = num;
                                    g_markers[g_markers.Count] = new DataPoint(Color.Black, new PointF(plot_x, highest_point), System.DateTime.Now);
                                }
                                else { /* old temp was higher */ } 
                                last_was_soak = true;
                                this_is_soak = false;
                            }
                            else
                            {  //no more soak step
                                this_is_soak = false;
                                last_was_soak = false;
                            }
                        }
                        #endregion

                        #region Negative Temperatures
                        //NEGATIVE TEMPERATURES
                        else
                        {
                            //CONVERTING FROM TWOS COMPLEMENT
                            //CONVERSION BYTE IS LESS THAN 128, IMPLICIT CONVERSION CAST
                            byte signedT = (byte)response[4];
                            if (signedT < -128)
                            {
                                string decade = (signedT / 10).ToString();
                                string s_right = signedT.ToString();

                                right = signedT.ToString().Substring(s_right.Length - 1, 1);
                                if (signedT < 99)
                                {
                                    tempLabel.Text = decade + "." + right;
                                }
                                else
                                {
                                    tempLabel.Text = signedT + "." + right;
                                }
                                neg = true;
                                current_y = signedT / 10;
                            }

                            //CONVERSION IS EXPLICIT, VALUES OVER 128 WILL OVERWRITE THE MSB WHICH HOLDS THE NEGATIVE BIT
                            if (signedT >= -128)
                            {
                                //the exact function as above, except values are subtracted from a reversed empty byte (255), and added to 12.8
                                //const_msb = short.Parse(response[3].ToString()) * 256; THE VALUE IS NOT EXPECTED TO PASS 128 DEGREES
                                short big_neg_num = (short)(128 + (128 - signedT)); //a formatted representation of 12.8 degrees plus actual twos compliment bits received
                                int big_divided_num = big_neg_num / 10;             // the whole part of the degrees
                                //int big_decimal_num = 128 - signedT;                // the decimal part of the degrees

                                string big_right = big_neg_num.ToString().Substring(big_neg_num.ToString().Length - 1, 1);

                                tempLabel.Text = "-" + big_divided_num + "." + big_right;
                                neg = true;
                                current_y = float.Parse(tempLabel.Text);// ~(big_neg_num / 10);
                            }

                            if (statusLabel.Text == "Soaking")
                            {
                                this_is_soak = true;
                                if ((this_is_soak == true) && (last_was_soak == false))
                                {
                                    g_markers.Add(null);
                                }

                                float num_neg = float.Parse(tempLabel.Text);
                                if (num_neg > lowest_point)
                                {
                                    lowest_point = num_neg;
                                    g_markers[g_markers.Count] = new DataPoint(Color.Black, new PointF(plot_x, highest_point), System.DateTime.Now);
                                }
                                else { }
                                last_was_soak = true;
                            }
                            else
                            {
                                this_is_soak = false;
                                last_was_soak = true; 
                            }
                        }

                        #endregion

                        panel2.Refresh();
                        tempLabel.Refresh();
                    }
                    catch { };
                    #endregion

                #region POST TO GUI
                    try
                    {
                        if (g_data.Count == 0)
                        {
                            plot_x = 1;
                            current_x = plot_x;
                            g_int = 0;
                            myG.Clear(Color.White);
                            panel1.BackgroundImage = WindowsFormsApplication1.Properties.Resources.graph_paper;
                            g_pt = new PointF(0, 0);
                            g_data.Add(new DataPoint(universal_color, g_pt,DateTime.Now));//starting point on graph
                        }

                        else
                        {
                            //CALCULATE THE CURRENT POINT
                            current_x = plot_x;
                            g_int += 1;
                            g_pt = new PointF(current_x, current_y);
                            g_data.Add(new DataPoint(universal_color, g_pt,DateTime.Now));

                            //DRAW THE TEMPERATURE POINTS
                            for (int i = 1; i < g_data.Count; i++)
                            {
                                    DataPoint curdatapoint = (DataPoint)g_data[i];
                                    myPen.Color = curdatapoint.calorie;

                                    DataPoint temppoint = (DataPoint)g_data[i-1];
                                    DataPoint olddot = new DataPoint(temppoint.calorie, temppoint.pit,DateTime.Now);

                                    float midpoint = panel1.Height / 2;
                                    float cury = curdatapoint.pit.Y / 60;
                                    float mappoint = midpoint - (cury * midpoint);
                                    float oldy = olddot.pit.Y / 60;
                                    float oldmappoint = midpoint - (oldy * midpoint);
                                    PointF p1 = new PointF(i / scale_x - 1, oldmappoint * scale_y);
                                    PointF p2 = new PointF(i / scale_x, mappoint * scale_y);

                                    myG.DrawLine(myPen, p1, p2);
                            }

                            // DRAW THE START/STOP MARKERS
                            foreach (DataPoint marker in g_markers)
                            {
                                float midpoint = panel1.Height / 2;
                                float cury = marker.pit.Y / 60;
                                float mappoint = midpoint - (cury * midpoint);

                                myG.DrawEllipse(new Pen(Brushes.ForestGreen, 5), new System.Drawing.Rectangle(new System.Drawing.Point((int)marker.pit.X / scale_x, (int)(mappoint *scale_y)), new Size(7, 7)));
                            }
                          }
                        }
                    catch (Exception) {  /* statusLabel.Text = "PlotData fail"; */}
                #endregion 

                            default_loop();
                    plot_x++;
                }
            else { timer1.Stop(); }
    }

        public void update_controls(int p_step)
        {
            if (curstep < 1) { curstep = 1; }

            if (!(currealprofile == null))
            {
                if (curstep > currealprofile.steps)
                {

                }
                else
                {
                        //Only way to GET and SET using cast ArrayList members
                        StepData localstepdata;
                        localstepdata = (StepData)currealprofile.data[curstep - 1];

                        if (localstepdata.type == "Ramprate" || localstepdata.type == "ramprate")
                        {
                            timebox.Text = " -- ";
                            setpointbox.Text = localstepdata.temp.ToString();
                            ratebox.Text = localstepdata.time.ToString();
                            if (editmode == true)
                            {
                                stepbox.Enabled = true;
                                timebox.Enabled = false;
                                setpointbox.Enabled = true;
                                ratebox.Enabled = true;
                            }
                        }

                        else if (localstepdata.type == "soak" || localstepdata.type == "Soak")
                        {
                            setpointbox.Text = " -- ";
                            ratebox.Text = " -- ";
                            timebox.Text = localstepdata.time.ToString();
                            if (editmode == true)
                            {
                                stepbox.Enabled = false;
                                setpointbox.Enabled = false;
                                ratebox.Enabled = false;
                                timebox.Enabled = true;
                            }
                        }
                        else if (localstepdata.type == "Ramptime" || localstepdata.type == "ramptime")
                        {
                            ratebox.Text = " -- ";
                            setpointbox.Text = localstepdata.temp.ToString();
                            timebox.Text = localstepdata.time.ToString();
                            if (editmode == true)
                            {
                                stepbox.Enabled = true;
                                timebox.Enabled = true;
                                setpointbox.Enabled = true;
                                ratebox.Enabled = false;
                            }
                        }
                        else if (localstepdata.type == "jump" || localstepdata.type == "Jump")
                        {
                            ratebox.Text = " -- ";
                            setpointbox.Text = " -- ";
                            timebox.Text = " -- ";
                            
                                stepbox.Enabled = false;
                                timebox.Enabled = false;
                                setpointbox.Enabled = false;
                                ratebox.Enabled = false;
                        }

                        else { setpointbox.Text = localstepdata.temp.ToString(); setpointbox.Enabled = true; }
                    }

                    if (editmode == false)
                    {
                        stepbox.Enabled = false;
                        timebox.Enabled = false;
                        ratebox.Enabled = false;
                        setpointbox.Enabled = false;
                    }
                }
            }

        public void calculate_time()
        {




        }

        void error_log(string msg, byte[] packet, string errmsg)
        {
            errl.WriteLine(msg + ":" + errmsg + Environment.NewLine);
            if (!(packet == null))
            {
                errl.WriteLine("ID: " + packet[0].ToString() + Environment.NewLine);
                errl.WriteLine("Fcode: " + packet[1].ToString() + Environment.NewLine);
                errl.WriteLine("Reg/Err: " + packet[2].ToString() + " " + packet[3].ToString() + Environment.NewLine);
                errl.WriteLine("Data: " + (packet[4]).ToString() + packet[5].ToString() + Environment.NewLine);
                errl.WriteLine("CRC: " + (packet[6]).ToString() + " " + packet[7].ToString() + Environment.NewLine);  //response strings are reverse order

                errl.WriteLine(Environment.NewLine);
            }
        }

        #region custom classes
        public class Profile
        {
            public string name;
            public int steps;
            public ArrayList data;

            public Profile() { }

            public Profile(string nam)
            {
                this.name = nam;
                this.data = new ArrayList(); // holds class StepData 
            }

            public Profile(string nam, StepData std)
            {
                this.name = nam;
                this.steps = 1;
                this.data = new ArrayList(); // holds class StepData 
                this.data.Add(std);
            }

        }

        public class StepData
        {
            public string type;
            public short temp;
            public byte time;
            public byte endvalue;

            public StepData() { }

            public StepData(string typ, byte val2)
            {
                //used to make soak step, temp"--"
                this.type = typ;
                this.time = val2;
            }

            public StepData(string typ, short val, byte val2)
            {
                this.type = typ;
                this.temp = val;
                this.time = val2;
            }

            public StepData(string typ, short val, byte val2, byte val3)
            {
                this.type = typ;
                this.temp = val;
                this.time = val2;
                this.endvalue = val3; //for end step post condition, and jump step repeat
            }
        }

        public class DataPoint
        {
            public Color calorie;
            public PointF pit;
            public DateTime dat;

            public DataPoint(DataPoint dp) { calorie = dp.calorie; pit = dp.pit;  }
            public DataPoint(Color cal, PointF pt, DateTime dt) { this.calorie = cal; this.pit = pt; this.dat = dt; }
            public DataPoint(System.Drawing.Point pt) { calorie = Color.Red; pit = pt; dat = System.DateTime.Now; }
        }
        #endregion

        #region Write Function
        public void WriteFunction(byte addy, byte func, ushort registers, short data)
        {
            if (sp.IsOpen)
            {
                //Clear in/out buffers:
                //sp.DiscardOutBuffer();
                sp.DiscardInBuffer();

                    BuildMessage(addy, func, registers, data, ref message);

                    //Send Modbus message to Serial Port:
                    try
                    {
                        try
                        {
                            sp.Write(message, 0, message.Length);
                            GetResponse(ref response);
                            error_log("Sent:", message, error_message);
                        }
                        catch(Exception merr) { errl.WriteLine("Error in Write:" + error_message + Environment.NewLine + merr + Environment.NewLine) ; }

                    }
                    catch (Exception)
                    {
                        //error_log("Sent:", message);
                        //error_log("No Response." , null);
                    }
                    //Evaluate message:
                    if (CheckResponse(response))
                    {
                        //error_log("Response:", response, error_message);
                    }
                    else
                    {
                        
                    }
            }
        }
        #endregion

        #region CRC Computation
        private void GetCRC(byte[] message, ref byte[] CRC)
        {
            //Function expects a modbus message of any length as well as a 2 byte CRC array in which to 
            //return the CRC values:

            ushort CRCFull = 0xFFFF;
            byte CRCHigh = 0xFF, CRCLow = 0xFF;
            char CRCLSB;

            for (int i = 0; i < (message.Length) - 2; i++)
            {
                CRCFull = (ushort)(CRCFull ^ message[i]);

                for (int j = 0; j < 8; j++)
                {
                    CRCLSB = (char)(CRCFull & 0x0001);
                    CRCFull = (ushort)((CRCFull >> 1) & 0x7FFF);

                    if (CRCLSB == 1)
                        CRCFull = (ushort)(CRCFull ^ 0xA001);
                }
            }
            CRC[1] = CRCHigh = (byte)((CRCFull >> 8) & 0xFF);
            CRC[0] = CRCLow = (byte)(CRCFull & 0xFF);
        }
        #endregion

        #region Build Message
        private void BuildMessage(byte address, byte function, ushort register, short data, ref byte[] message)
        {
            //Array to receive CRC bytes:
            byte[] CRC = new byte[2];

            message[0] = address;
            message[1] = (byte)(function);
            message[2] = (byte)(register >> 8);
            message[3] = (byte)(register);
            message[4] = (byte)(data >> 8); //number of registers to read
            message[5] = (byte)(data);
            GetCRC(message, ref CRC);

            message[message.Length - 2] = CRC[0];
            message[message.Length - 1] = CRC[1];
        }
        #endregion

        #region Check Response
        private bool CheckResponse(byte[] response)
        {
            //Perform a basic CRC check:
            byte[] CRC = new byte[2];
            GetCRC(response, ref CRC);
            if (CRC[0] == response[response.Length - 2] && CRC[1] == response[response.Length - 1])
                return true;
            else
                return false;
        }
        #endregion

        #region Get Response
        private void GetResponse(ref byte[] response)
        {
            sp.ReadTimeout = 500;
            for (int i = 0; i < response.Length; i++)
            {
                response[i] = (byte)(sp.ReadByte());
            }
        }
        #endregion

        private void increaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myPen.Width++;
        }

        private void decreaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myPen.Width--;
        }

        public void EstablishConnection()
        {
            #region open serial port
            if (!(sp.IsOpen))
            {
                try
                {
                    sp.Open();
                }
                catch (Exception)
                {
                    try
                    {
                        sp = new SerialPort("COM4", 9600);
                        sp.Open();
                    }
                    catch (Exception)
                    {

                        try
                        {
                            sp = new SerialPort("COM4", 9600);
                            sp.Open();
                        }
                        catch (Exception)
                        {
                            try
                            {
                                sp = new SerialPort("COM6", 9600);
                                sp.Open();
                            }
                            catch (Exception)
                            {
                                checkBox1.Checked = false;
                                MessageBox.Show("COM4,5,6 not ready", "MsgBox",
                                MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                            }
                        }
                    }
                }
            }
            else
            {
                //errl.WriteLine("using" + sp.PortName + Environment.NewLine);
            }
                #endregion

        }

        private void commsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sp.IsOpen == false)
            {
                try
                {
                    EstablishConnection();
                    modbusForm commsform = new modbusForm(sp);
                    commsform.Show();
                }
                catch (Exception) { };
            }
            else
            {
                modbusForm commsform = new modbusForm(sp);
                commsform.Show();
            }
        }

        private void readToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sp.IsOpen == false)
                try{EstablishConnection();} catch(Exception){}
            if (sp.IsOpen == true)
            {
                EstablishConnection();

                System.Windows.Forms.TextBox output = new System.Windows.Forms.TextBox();
                output.Multiline = true;
                output.Width = this.Width;
                output.Height = this.Height;
                Form textwindow = new Form();
                textwindow.Controls.Add(output);

                WriteFunction(01, 03, 102, 01);     //get alarm1 status 
                string resp1 = response[4] == 0 ? "OK" : "Alarm Tripped";
                WriteFunction(01, 03, 702, 01);     //get alarm1 type
                string resp2 = response[4].ToString();
                WriteFunction(01, 03, 700, 01);     //get alarm1 source
                string resp3 = response[4] == 1 ? "Heat" : "Cool";

                output.AppendText("Response: " + Environment.NewLine +
                                   "Alarm1 Status:    " + resp1  + Environment.NewLine +
                                   "Alarm1 Type:  " + resp2 + "      <Process, Deviation, MaxRate>" + Environment.NewLine +
                                   "Monitor Function:   " + resp3  + Environment.NewLine);

                WriteFunction(01, 03, 106, 01);     //get alarm1 status 
                string resp1b = response[4] == 0 ? "OK" : "BAD";
                WriteFunction(01, 03, 719, 01);     //get alarm1 type
                string resp2b = response[4].ToString();
                WriteFunction(01, 03, 717, 01);     //get alarm1 source
                string resp3b = response[4] == 1 ? "Heat" : "Cool";

                output.AppendText(Environment.NewLine +
                                   "Alarm2 Status:   " + resp1b + Environment.NewLine +
                                   "Alarm2 Type:  " + resp2b + "      <Process, Deviation, MaxRate>" + Environment.NewLine +
                                   "Monitor Function:   " + resp3b + Environment.NewLine);
                textwindow.Show();
            }
            else 
            {
                MessageBox.Show("Error", "No Connection."); 
            }
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WriteFunction(01, 06, 331, 0); // Alarm2 clear heating
            WriteFunction(01, 06, 102, 0); // Alarm2 clear heating

            WriteFunction(01, 06, 312, 0); // Alarm1 clear cooling
            WriteFunction(01, 06, 106, 0); // Alarm1 clear cooling

            WriteFunction(01, 06, 704, 0); // alarm latch clear 1
            WriteFunction(01, 06, 721, 0); // alarm latch clear 2
            WriteFunction(01, 06, 607, 0); // error latch clear 1
            WriteFunction(01, 06, 617, 0); // error latch clear 2
        }

        private void temperatureLimitsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int locX = MousePosition.X;
            int locY = MousePosition.Y;
            string responseHI = "65";
            responseHI = Microsoft.VisualBasic.Interaction.InputBox("Enter HIGH Temperature Limits (in Celcius)", "Limits", responseHI, locX + 50, locY);
            string responseLOW = "-35";
            responseLOW = Microsoft.VisualBasic.Interaction.InputBox("Enter LOW Temperature Limits (in Celcius)", "Limits", responseLOW, locX + 50, locY);

            if (sp.IsOpen == false)
            {
                EstablishConnection();
            }
                try
                {
                    short high = short.Parse(responseHI.ToString());
                    WriteFunction(01, 06, 303, (short)(high * 10));
                    if (!(response[2] == 1))
                    {
                        MessageBox.Show("Set Hight limit command failed!", "MsgBox",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                    }

                    short low = short.Parse(responseLOW.ToString());
                    WriteFunction(01, 06, 302, (short)(low * 10));
                    if (!(response[2] == 1))
                    {
                        MessageBox.Show("Set Low limit command failed!", "MsgBox",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                    }
                }
                catch (Exception) 
                { 
                //happens if the user hits the cancel button 
                }
            }

        private void jumpStepToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 jumpform = new Form2(Form1.ActiveForm);
            jumpform.Show();
        }

        private void passwordToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (sp.IsOpen == false) { try { EstablishConnection();} catch(Exception ){}}
            if (sp.IsOpen == true)
            {
                string pw = Microsoft.VisualBasic.Interaction.InputBox("Enter new password (4 chars)", "Set Password", "PASS", MousePosition.X + 50, MousePosition.Y);

                byte[] password = new byte[4];
                char[] part = pw.ToCharArray();
                password[0] = (byte)part[0];
                password[1] = (byte)part[1];
                password[2] = (byte)part[2];
                password[3] = (byte)part[3];


                WriteFunction(01, 06, 1314, 1); //Set the bit.
                ushort i = 1330;
                foreach (byte bee in password)
                {
                    WriteFunction(01, 06, i, bee);
                    i++;
                }
                WriteFunction(01, 06, 1314, 0); //Clear the bit.

                MessageBox.Show("PASSWORD SUCCESSFUL","PW Set");
            }
            else { }

        }

        private void saveAndQuitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // add functionality to capture profile name as read from TreeNode.name
            string myname = currealprofile.name;
            TextWriter tw = new StreamWriter(myname + ".txt");

            tw.WriteLine(currealprofile.name);
            tw.WriteLine(currealprofile.steps);

            foreach (StepData itemd in currealprofile.data)
            {
                tw.WriteLine("[" + itemd.type + "," + itemd.temp.ToString() + "," + itemd.time.ToString() + "," + itemd.endvalue.ToString() + "]");
            }
            tw.Close();
            statusLabel.Text = "saved.";
        }

        private void dontSaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { sp.Close(); }
            catch (Exception) { }
            this.Close();
        }

        private void resetVerticalToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            scale_y = 1;
        }

        private void resetPenSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myPen.Width = 3;
        }

        private void increaseVerticalToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            scale_y = scale_y * 1.25f;
        }

        private void decreaseVerticalToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (scale_y == 0) { }
            else
            {
                scale_y = scale_y * .8f;
            }
        }

        private void lockPanelsToolStripMenuItem_Click(object sender, EventArgs e)
        {
           panel3.Enabled = panel3.Enabled == true ? false : true;
        }

        private void exportExcelDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];
           
            ws.Cells[1, 1] = System.DateTime.Now;
           
            ((Range)ws.Cells[1, 1]).EntireColumn.AutoFit();
            ((Range)ws.Cells[3, 3]).EntireColumn.ColumnWidth = 22;

            int k = 3;
               //Range newrange = ws.Cells.get_Range(1, g_data.Count);
               foreach (DataPoint data in g_data)
               {
                   ws.Cells[k, 1] = data.pit.X;
                   ws.Cells[k, 2] = data.pit.Y;
                   ws.Cells[k, 3] = data.dat;
                   k++;
               }
               ChartObjects standardplot = (ChartObjects)ws.ChartObjects(Type.Missing);
               ChartObject Plotter = standardplot.Add(300, 100, 500, 250);
               Chart xlchart = Plotter.Chart;
                xlchart.SetSourceData(((Range)ws.Cells[3,2]).EntireColumn,XlRowCol.xlColumns);
                xlchart.ChartType = XlChartType.xlXYScatterLines;
                xlchart.AutoScaling = true;
            
                Axis xAxis = xlchart.Axes(XlAxisType.xlCategory);
                xAxis.TickLabels.Orientation = XlTickLabelOrientation.xlTickLabelOrientationUpward;
                
                //ws.ChartObjects(xlchart).Activate();
                //xlchart.Axes(xlValue).MajorGridlines.Select();
                xlchart.SeriesCollection(1).XValues = ((Range)ws.Cells[3,3]).EntireColumn;
                xAxis.TickLabels.NumberFormat = "[$-409]h:mm AM/PM";

                xAxis = (Axis)xlchart.Axes(XlAxisType.xlCategory);
                xAxis.HasTitle = true;
                xAxis.AxisTitle.Text = "Time";
                
                Axis yAxis = (Axis)xlchart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                yAxis.HasTitle = true;
                yAxis.AxisTitle.Text = "Temp";

                xlchart.HasTitle = true;
                xlchart.ChartTitle.Text = "Thermal Test Record " + DateTime.Now.DayOfWeek + " " + DateTime.Now.Month +"-"+ DateTime.Now.Day + "-" + DateTime.Now.Year ;
                xlchart.HasLegend = false;

        }

        private void intervalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s_time;
            int time;
            
           s_time =  Microsoft.VisualBasic.Interaction.InputBox("Seconds: ", "Reporting Interval", "5");
           try
           {
               time = int.Parse(s_time);
               if (time > 5)
               {
                   App_interval = time * 1000;
               }
           }
           catch (Exception) { }
        }


    }
}
