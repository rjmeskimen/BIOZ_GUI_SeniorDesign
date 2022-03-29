using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;

namespace UART_Senior_Design_Test
{
    public partial class Bluetooth_Settings : Form
    {
        public SerialPort _serial = new SerialPort();                           //create an instance of the serial port
        public SerialPort _serial2 = new SerialPort();                          //create another instnace
        public Bluetooth_Settings()
        {
            InitializeComponent();
            _serial.BaudRate = int.Parse(baud_rate.Text);                       //set the baud rate of the serial ports to the int value of the string in the drop down box
            _serial2.BaudRate = int.Parse(baud_rate.Text);
            foreach (string s in SerialPort.GetPortNames())
            {
                com_port.Items.Add(s);                                          //get all the open ports from the serial port and populate the drop down box
                com_port2.Items.Add(s);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Connect_Click(object sender, EventArgs e)                  //fire this event when the Connect Button is CLicked
        {
            try
            {
                _serial.PortName = com_port.SelectedItem.ToString();            //set the first comport to the selected item from the combo box
                _serial2.PortName = com_port2.Text.ToString();

                _serial.BaudRate = Convert.ToInt32(baud_rate.SelectedItem);
                _serial2.BaudRate = Convert.ToInt32(baud_rate.SelectedItem);    //set the serial port baud rate to the selected combo box

                _serial.Open();                                                 //open the serial port

                if (_serial2.PortName == "COM99")
                {

                }
                else
                {
                    _serial2.Open();
                }
                

               this.Hide();                                                   //Hide this Form
                Form1 _main = new Form1();                                      //create instance of the Form1
                foreach (Form1 tmpform in Application.OpenForms)
                {
                    if(tmpform.Name == "Form1")
                    {
                        _main = tmpform;
                        break;
                    }
                }
                _main.toolStripStatusLabel1.Text = " Connected: " + _serial.PortName.ToString() + "as Port1 & " + _serial2.PortName.ToString() + "as Port2";
                _main.toolStripStatusLabel1.ForeColor = Color.Green;
                _main.toolStripProgressBar1.Value = 100;
            }
            catch
            {
                MessageBox.Show("Please select proper COM Port/Baud Rate");
            }

        }

        private void button1_Click(object sender, EventArgs e)                                 //Fire this event when the AUTO CONNECT button is selected
        {

            string use_com = "COM4";                                                            //*******TODO: Edit these strings to declare which com ports we should auto connect to
            string use_com2 = "COM5";
            try
            {
                _serial.PortName = use_com;
                _serial2.PortName = use_com2;

                _serial.BaudRate = Convert.ToInt32(baud_rate.SelectedItem);
                _serial2.BaudRate = Convert.ToInt32(baud_rate.SelectedItem);

                _serial.Open();

                try
                {
                    _serial2.Open();
                }
                catch
                {
                    MessageBox.Show("Only " + _serial.PortName.ToString() +" was connected." + _serial2.PortName.ToString() + "was not connected. Proceed with only one serial port connection");
                }


                this.Close();

                Form1 _main = new Form1();
                foreach (Form1 tmpform in Application.OpenForms)
                {
                    if (tmpform.Name == "Form1")
                    {
                        _main = tmpform;
                        break;
                    }
                }

                _main.toolStripStatusLabel1.Text = " Connected: " + _serial.PortName.ToString() + "as Port1 & " + _serial2.PortName.ToString() + "as Port2";
                _main.toolStripStatusLabel1.ForeColor = Color.Green;
                _main.toolStripProgressBar1.Value = 100;
            }
            catch
            {
                MessageBox.Show("Please select proper COM Port/Baud Rate to Auto Connect");
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {

            string use_com = "COM1";                                                            //*******TODO: Edit these strings to declare which com ports we should auto connect to
            try
            {
                _serial.PortName = use_com;

                _serial.BaudRate = Convert.ToInt32(38400);

                _serial.Open();

                this.Close();

                Form1 _main = new Form1();
                foreach (Form1 tmpform in Application.OpenForms)
                {
                    if (tmpform.Name == "Form1")
                    {
                        _main = tmpform;
                        break;
                    }
                }

                _main.toolStripStatusLabel1.Text = " Connected: " + _serial.PortName.ToString() + "as Port1 & ";
                _main.toolStripStatusLabel1.ForeColor = Color.Purple;
                _main.toolStripProgressBar1.Value = 100;
            }
            catch
            {
                MessageBox.Show("Please select proper COM Port/Baud Rate to Auto Connect");
            }
        }

        private void baud_rate_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
