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
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Media;
using NPOI.SS.UserModel;
using ZedGraph;

//UUUSSING


namespace UART_Senior_Design_Test
{


    public partial class Form1 : Form
    {
        bool start_parsing = false;
        int value, value2;
        double updatevalue;
        double decimal_updatevalue;

        static int xl_width = 18;
        static int xl_length = 56;
        int num_of_sheets = 2;

        double[] Data_Array = new double[24];
        double[,] data_array = new double[(xl_length * 24) + 1, xl_width];
        double[,] data_array1 = new double[xl_length, xl_width];
        double[,] data_array2 = new double[xl_length, xl_width];
        double[,] data_array3 = new double[xl_length, xl_width];

        int r = 1, c = 0, sheet_count = 0;
        bool decimal_mode, negative_mode = false;
        bool start_recording = false;
        int j = 0; 
        char switch1, switch2, switch3, switch4;
        int btn_count = 0;

        bool busy = false;

        private Bluetooth_Settings _setting = new Bluetooth_Settings();                        //create an instance of the Bluetooth_Settings
        private Form4 _helpmenue = new Form4();
        public string frequency;
        double dacgain = 1;
        double exbuffgain = 2;
        char dacgaincode, exbuffgaincode = '0';
        double finalpk2pk = 300;



        GraphPane graphPane;

        
        public Form1()                                                                         //Initialize the form
        {
            InitializeComponent();
            STOCK_img.BringToFront();
            bring_labels_to_front();
            richTextBox2.Text = Convert.ToString(offset_scroll.Value);
            pk2pk_textbox.Text = "800";
            //graphPane = zedGraphControl1.GraphPane;

            finalpk2pk = Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain;
            final_pk2pk_textbox.Text = Convert.ToString(finalpk2pk);
            DrawSine();


        }

        double excitationOffset = 1100;

        


        private void DrawSine()
        {
            graphPane = zedGraphControl1.GraphPane;
            PointPairList _pointPairList = new PointPairList();

            _pointPairList.Clear();
            for (int _angle = 0; _angle <= 360; _angle = _angle +10)
            {
                double _x = _angle;
                double _y = (((finalpk2pk /2) * (Math.Sin(Math.PI * _x / 180.0)) + excitationOffset)) ;
                double _y2 = (Math.Sin(Math.PI * _x / 180.0));

                if (_y >= 2400)     //if saturates at 2.4V
                {
                    _y = 2400;
                }
                else if (_y <= 200)
                {
                    _y = 200;   
                }
                else
                {
                    //y does not saturate
                }
                
                PointPair _pointPair = new PointPair(_x, _y);

                _pointPairList.Add(_pointPair);
                zedGraphControl1.Invalidate();
            }

            {
                //double _x
                //double _y2 =
                //PointPair _pointPairMax = new PointPair(_x, _y);
            }

                graphPane.CurveList.Clear();
                LineItem _lineItem = graphPane.AddCurve("Excitation Output", _pointPairList, Color.Red, SymbolType.None);
            //LineItem _lineItem_max = graphPane.AddCurve("Excitation Output", _pointPairList_maxline, Color.Red, SymbolType.None);
            //LineItem _lineItem_min = graphPane.AddCurve("Excitation Output", _pointPairList_minline, Color.Red, SymbolType.None);

            graphPane.Title.Text = "Excitation Output";                 //set the title and axis'
                //graphPane.XAxis.Title.Text = "Time (ms)";
                graphPane.YAxis.Title.Text = "Voltage (mV)";
                zedGraphControl1.AxisChange();
                zedGraphControl1.PerformAutoScale();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void bluetoothSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void cOMPortToolStripMenuItem_Click(object sender, EventArgs e)                //open tool strip form
        {
            //Form 2 Bluetooth Display
            _setting.Show();                                                                                    //open up the tool strip form
            _setting._serial.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);           //when data has been recieved by Serial Port 1 -> add the event 'DataRecievedHandler'
            _setting._serial2.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler2);         //when data has been recieved by Serial Port 2 -> add the event 'DataRecievedHandler2'
        }
        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)         //data recieved handler for Serial POrt1. Put into array and call Extract each character
        {
            string indata;
            SerialPort sp = (SerialPort)sender;
            while (sp.ReadBufferSize != 0)
            {
                indata = sp.ReadLine();
                this.Invoke(new EventHandler(update_richtextbox1), new object[] { indata });
                this.Invoke(new EventHandler(Parse_My_Data), new object[] { indata });
            }

        }
        private void Parse_My_Data(object sender, EventArgs e)                                 //parse the incoming data into an array
        {
            char data;
            string indata = (string)sender;
            for (int i = 0; i < indata.Length; i++)
            {
                data = indata[i];
                if (data == '*')
                {
                    start_recording = true;
                    Array.Clear(data_array, 0, data_array.Length);
                    r = 1;
                    c = 0;
                }

                if (start_recording)
                {
                    if (data >= 48 && data <= 57 || data == '.' || data == '-')                      //if the byte passed in through the serial port is a number
                    {
                        if (decimal_mode)
                        {
                            int data2 = (int)data - 48;                                              //convert the numerical value
                            value2 = (data2);                                                        //set value to the numerical data
                            decimal_updatevalue = decimal_updatevalue * 10 + value2;                 //get a final value based on how many int were entered before a ',' is recognized
                            value2 = 0;
                            j++;
                            //reset value
                        }

                        else if (data == '.')
                        {
                            decimal_mode = true;                                                     //notify that the following numbers will be decimals until a comma is detected
                            j = 0;
                        }
                        else if (data == '-')
                        {
                            negative_mode = true;
                        }
                        else
                        {
                            int data1 = (int)data - 48;                                              //convert the numerical value
                            value = (data1);                                                         //set value to the numerical data
                            updatevalue = updatevalue * 10 + value;                                  //get a final value based on how many int were entered before a ',' is recognized
                            value = 0;                                                               //reset value
                        }
                    }
                    if (data == ',')                                                                 //once the , is recognized. save that data into the array
                    {
                        if (decimal_mode)
                        {
                            double temp = decimal_updatevalue / (Math.Pow(10, j));
                            updatevalue = updatevalue + temp;
                        }
                        if (negative_mode)
                        {
                            updatevalue = updatevalue * -1;
                        }

                        data_array[r, c] = updatevalue;
                        c++;                                                                       //increment the column

                        if (c == xl_width - 6)
                        {
                            r++;                                                                   //after 8 column datas recoded, then increment the row
                            c = 0;                                                                 //reset the column back to 0;
                            sheet_count++;
                        }
                        if (sheet_count == xl_length-1)
                        {
                            sweep_number_label.Text = Convert.ToString(num_of_sheets);
                            num_of_sheets++;
                            SystemSounds.Beep.Play();
                            sheet_count = 0;

                            switch (num_of_sheets-1)
                            {
                                case 1:
                                case 4:
                                    V01.BringToFront();
                                    break;
                                case 2:
                                case 7:
                                    V02.BringToFront();
                                    break;
                                case 3:
                                case 10:
                                    V03.BringToFront();
                                    break;                                
                                case 5:
                                case 8:
                                    V12.BringToFront();
                                    break;                                
                                case 11:
                                case 6:
                                    V31.BringToFront();
                                    break;
                                case 12:
                                case 9:
                                    V32.BringToFront();
                                    break;
                            }

                        }


                        updatevalue = 0;                                                           //since a ',' was recognized. reset the update value
                        decimal_updatevalue = 0;
                        value2 = 0;
                        value = 0;
                        decimal_mode = false;                                                      //set low so the next number is not recorded as a decimal
                        negative_mode = false;
                        j = 0;

                    }
                    
                }
            }

        }
        private void DataReceivedHandler2(object sender, SerialDataReceivedEventArgs e)        //data recieved handler for Serial Port2
        {
            string indata;
            SerialPort sp = (SerialPort)sender;
            while (sp.ReadBufferSize != 0)
            {
                indata = sp.ReadLine();
                this.Invoke(new EventHandler(remote_handler), new object[] { indata });
            }
        }
        private void remote_handler(object sender, EventArgs e)
        {
            int frequency = 0;
            int sweepRange = 0;
            string data = (string)sender;

            char command = data[0];

            data = data.Remove(0, 1);
            if(busy)
            {
                sendBusyRemote();
            }
            if (command == 'F') //Single Frequency
            {
                frequency = Convert.ToInt32(data);
                //TakeOneSamplePort2(data);
                //_setting._serial2.Write("We got " + Convert.ToString(frequency) + "#");

            }
            else if(command == 'S') //Sweep Frequency
            {
                sweepRange = Convert.ToInt32(data);
                sendSweep(sweepRange);
            }
            else if(command == 'E') //Entire Range
            {
                sendEntirerange();
            }
            else if(command == 'C')
            {
                richTextBox1.Clear();
            }
        }
        public void TakeOneSamplePort2(string frequencyPassedIn)
        {
            string freq = frequencyPassedIn;

            char[] sendToChip = new char[1];
            //frequency = richTextBox4.Text;

            //richTextBox1.AppendText(Convert.ToString(richTextBox4.Text[0]));

            sendToChip[0] = 'F';
            _setting._serial.Write(sendToChip, 0, 1);

            quickPause();

            for (int i = 0; i < freq.Length; i++)
            {
                sendToChip[0] = Convert.ToChar(freq[i]);
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            sendToChip[0] = 'F';
            _setting._serial.Write(sendToChip, 0, 1);
            quickPause();

            sendToChip[0] = 'H';
            _setting._serial.Write(sendToChip, 0, 1);
        }
        public void sendSweep(int sweepType)
        {
            char[] sendToChip = new char[1];

            switch (sweepType)
            {
                case 1://"1 - 200k Hz"
                    {
                        sendToChip[0] = '!';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                case 2: //"1 - 10 Hz"
                    {
                        sendToChip[0] = '@';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                case 3: //"10 - 100 Hz"
                    {
                        sendToChip[0] = '#';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                case 4: //"100 - 1k Hz"
                    {
                        sendToChip[0] = '(';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                case 5: //"1k - 10k Hz"
                    {
                        sendToChip[0] = '%';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                case 6: //"10k - 100k Hz"
                    {
                        sendToChip[0] = '^';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                case 7: //"100k - 200k Hz"
                    {
                        sendToChip[0] = '&';
                        _setting._serial.Write(sendToChip, 0, 1);
                        quickPause();
                        break;
                    }
                default:
                    {
                        MessageBox.Show("A Sweep Was Not Sent To MCU");
                        break;
                    }
            }
        }

        private void sendBusyRemote()
        {
            char[] sendToChip = new char[1];
            sendToChip[0] = '!';
            _setting._serial2.Write(sendToChip, 0, 1);
        }

        private void sendEntirerange()
        {
            char[] sendToChip = new char[1];
            sendToChip[0] = 'E';
            _setting._serial.Write(sendToChip, 0, 1);
            quickPause();
        }

        private void Extract_each_Character(byte[] buffer, int count)                          //extract characters from the 1 serial port and proccess the data as needed
        {
            byte data;
            for (int i = 0; i < count; i++)
            {
                data = buffer[i];
                //Selects 
                this.Invoke(new EventHandler(update_richtextbox1), new object[] { data });
                this.Invoke(new EventHandler(Parse_My_Data), new object[] { data });
            }
        }
        
        private void update_richtextbox6(object sender, EventArgs e)                           //update when Data is recieved from Serial Port2
        {
            //  richTextBox6.
            byte data = (byte)sender;
            char s = Convert.ToChar(data);
            richTextBox6.Focus();
            richTextBox6.TextChanged -= new EventHandler(richTextBox6_TextChanged);
            richTextBox6.AppendText(s.ToString());
            richTextBox6.TextChanged += new EventHandler(richTextBox6_TextChanged);
        }
        private void update_richtextbox1(object sender, EventArgs e)                           //update when Data is recieved from Serial Port1
        {
            //  richTextBox1.
            // byte data = (byte)sender;
            //char s = Convert.ToChar(data);
            richTextBox1.Focus();
            richTextBox1.TextChanged -= new EventHandler(richTextBox1_TextChanged);
            richTextBox1.AppendText((string)sender + "\n");
            richTextBox1.TextChanged += new EventHandler(richTextBox1_TextChanged);
        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)                      //Update this text box when data is recieved from serial port1
        {
            int text_length = 0;
            text_length = richTextBox1.TextLength;
            if (text_length != 0)
            {
                char send_ch = richTextBox1.Text[text_length - 1];
                char[] ch = new char[1];
                ch[0] = send_ch;
                //richTextBox1.Undo();
                _setting._serial.Write(ch, 0, 1);
            }

        }
        private void richTextBox6_TextChanged(object sender, EventArgs e)                      //Update this text box when data is recieved from serial port2
        {
            int text_length = 0;
            text_length = richTextBox1.TextLength;
            if (text_length != 0)
            {
                char send_ch = richTextBox1.Text[text_length - 1];
                char[] ch = new char[1];
                ch[0] = send_ch;
                //richTextBox1.Undo();
                _setting._serial.Write(ch, 0, 1);
            }
        }
        private void clearRTB1_Click(object sender, EventArgs e)
        {
            //_setting._serial.DiscardInBuffer();
            richTextBox1.Clear();
            // SystemSounds.Asterisk.Play();
            SystemSounds.Beep.Play();
            // SystemSounds.Hand.Play();
            //SystemSounds.Question.Play();

        }                            //Clear the textbox
        private void playSimpleSound()
        {
            //SoundPlayer simpleSound = new SoundPlayer("need to enter location of a .wav file");

        }                                                       //play a sound
        private void button1_Click(object sender, EventArgs e)                                  //Take one sample. Record the Freq and other ADC Parameters
        {

            TakeOneSample();                                                                 //Take one sample. Record the Freq and other ADC Parameters

        }
        private void quickPause()
        {
            for (int i = 0; i < 1000; i++)
            {

            }
        }                                                            //for loop pause
        private void button2_Click_1(object sender, EventArgs e)                                //Address Read button click
        {
            char[] sendToChip = new char[1];

            sendToChip[0] = 'R';
            _setting._serial.Write(sendToChip, 0, 1);

            quickPause();

            for (int i = 0; i < richTextBox5.TextLength; i++)
            {
                sendToChip[0] = Convert.ToChar(richTextBox5.Text[i]);
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            sendToChip[0] = 'R';
            _setting._serial.Write(sendToChip, 0, 1);
            quickPause();
        }
        private void label7_Click(object sender, EventArgs e)
        { }                                //nothing
        private void richTextBox7_TextChanged(object sender, EventArgs e)
        { }                    //nothing
        private void richTextBox9_TextChanged(object sender, EventArgs e)
        { }                    //nothing
        private void richTextBox8_TextChanged(object sender, EventArgs e)
        { }                    //nothing
        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {
            richTextBox8.Text = Convert.ToString((800000 / Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text)));
            if (richTextBox4.Text == "")
            { }
            else
            {
                richTextBox9.Text = Convert.ToString((Convert.ToInt64(comboBox3.Text) * Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text) * Convert.ToInt64(richTextBox4.Text) / 800000));       //Cycles = DFTnum * sinc2 * sinc3 * excitation / ADC_Clk
                if (Convert.ToInt64(richTextBox9.Text) == 0)
                {
                    richTextBox7.Text = " ";
                }
                else
                {
                    richTextBox7.Text = Convert.ToString((Convert.ToInt64(comboBox3.Text) / Convert.ToInt64(richTextBox9.Text)));                //samples/cycle = DFTNum / total cycles

                }
            }

        }                    //calculate the samples/cycle
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)                 //Update the cycles & samples/cycle
        {
            richTextBox8.Text = Convert.ToString((800000 / Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text)));
            richTextBox9.Text = Convert.ToString(((Convert.ToInt64(comboBox3.Text) * Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text) * Convert.ToInt64(richTextBox4.Text) / 800000)));       //Cycles = DFTnum * sinc2 * sinc3 * excitation / ADC_Clk
            if (Convert.ToInt64(richTextBox9.Text) == 0)
            {
                richTextBox7.Text = " ";
            }
            else
            {
                richTextBox7.Text = Convert.ToString((Convert.ToInt64(comboBox3.Text) / Convert.ToInt64(richTextBox9.Text)));                //samples/cycle = DFTNum / total cycles

            }
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)                 //Update the cycles & samples/cycle
        {
            richTextBox8.Text = Convert.ToString((800000 / Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text)));
            richTextBox9.Text = Convert.ToString(((Convert.ToInt64(comboBox3.Text) * Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text) * Convert.ToInt64(richTextBox4.Text) / 800000)));       //Cycles = DFTnum * sinc2 * sinc3 * excitation / ADC_Clk
            if (Convert.ToInt64(richTextBox9.Text) == 0)
            {
                richTextBox7.Text = " ";
            }
            else
            {
                richTextBox7.Text = Convert.ToString((Convert.ToInt64(comboBox3.Text) / Convert.ToInt64(richTextBox9.Text)));                //samples/cycle = DFTNum / total cycles

            }
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)                //Update the cycles & samples/cycle
        {
            richTextBox8.Text = Convert.ToString((800000 / Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text)));
            richTextBox9.Text = Convert.ToString(((Convert.ToInt64(comboBox3.Text) * Convert.ToInt64(comboBox4.Text) * Convert.ToInt64(comboBox5.Text) * Convert.ToInt64(richTextBox4.Text) / 800000)));       //Cycles = DFTnum * sinc2 * sinc3 * excitation / ADC_Clk
            if (Convert.ToInt64(richTextBox9.Text) == 0)
            {
                richTextBox7.Text = " ";
            }
            else
            {
                richTextBox7.Text = Convert.ToString((Convert.ToInt64(comboBox3.Text) / Convert.ToInt64(richTextBox9.Text)));                //samples/cycle = DFTNum / total cycles

            }
        }
        static int DFT_TABLE = 0;
        static int SINC2_TABLE = 1;
        static int SINC3_TABLE =2;
        static int DFT_SRC =3;
        static int PWR_MOD =4;
        static int VOLT_MAG =5;
        static int VOLT_PHASE =6;
        static int CURRENT_MAG =7;
        static int CURRENT_PHASE =8;
        static int FREQ_RESULT =9;
        static int MAG_RESULT =10;
        static int PHASE_RESULT =11;
        static int DFT_VALUE =12;
        static int SINC2_VALUE =13;
        static int SINC3_VALUE =14;
        static int CYCLES =15;
        static int SAMP_CYCLE =16;
        static int SAMP_FREQ =17;

        private void SaveData()
        {
            //data_array[1400, 11];
            //C0 = dftnum table
            //C1 = sinc2 table
            //C2 = sinc3 table
            //C3 = dft src
            //C4 = pwr mod
            //C5 = Volt Mag
            //C6 = Volt phase
            //C7 = Curr Mag
            //C8 = Curr phase
            //C9 = freq result
            //C10 = Mag result
            //C11 = phase result
            //C12 = translated dft value
            //C13 = translated sinc2 value
            //C14= translated sinc3 value
            //C15= calculated cycles = dft*sinc2*sinc3 / adc Clk
            //C16= samples/cycle = dft / cycles
            //c17= sampling freq = adc clk / (sinc2*sinc3)
            
            

            for (int a = 1; a < xl_length * 24 + 1; a++)                                                                              //assign column 11 to the sinc2 value
            {
                switch (data_array[a, SINC2_TABLE])
                {
                    case 0:
                        data_array[a, SINC2_VALUE] = 1;
                        break;
                    case 1:
                        data_array[a, SINC2_VALUE] = 22;
                        break;
                    case 2:
                        data_array[a, SINC2_VALUE] = 44;
                        break;
                    case 3:
                        data_array[a, SINC2_VALUE] = 89;
                        break;
                    case 4:
                        data_array[a, SINC2_VALUE] = 178;
                        break;
                    case 5:
                        data_array[a, SINC2_VALUE] = 267;
                        break;
                    case 6:
                        data_array[a, SINC2_VALUE] = 533;
                        break;
                    case 7:
                        data_array[a, SINC2_VALUE] = 640;
                        break;
                    case 8:
                        data_array[a, SINC2_VALUE] = 667;
                        break;
                    case 9:
                        data_array[a, SINC2_VALUE] = 800;
                        break;
                    case 10:
                        data_array[a, SINC2_VALUE] = 889;
                        break;
                    case 11:
                        data_array[a, SINC2_VALUE] = 1067;
                        break;
                    case 12:
                        data_array[a, SINC2_VALUE] = 1333;
                        break;
                    default:
                        data_array[a, SINC2_VALUE] = 0;
                        break;
                }
                //Now do the same for the sinc3
                switch (data_array[a, SINC3_TABLE])
                {
                    case 0:
                        data_array[a, SINC3_VALUE] = 2;
                        break;
                    case 1:
                        data_array[a, SINC3_VALUE] = 4;
                        break;
                    case 2:
                        data_array[a, SINC3_VALUE] = 5;
                        break;
                    default:
                        data_array[a, SINC3_VALUE] = 0;
                        break;
                }
            }                                                           //assign column 9 to the sinc2 value & C10 to the sinc3 value

            for (int a = 1; a < xl_length * 24 + 1; a++)                                           //assign column 12 to the translated DFT value
            {
                //double dftnum_example = 4 * Math.Exp(.6931475*dft_table);
                data_array[a, DFT_VALUE] = 4 * Math.Exp(.6931475 * data_array[a, DFT_TABLE]);                // DFT value  = 4*e^.6931475*dft   ::: start at row 0, column 12
            }                                                     //assign column 12 to the translated DFT value

            for (int a = 1; a < xl_length * 24 + 1; a++)
            {
                data_array[a, 15] = data_array[a, SINC2_VALUE] * data_array[a, SINC3_VALUE] * data_array[a, DFT_VALUE] * data_array[a, FREQ_RESULT] / 800000;       //calculate total cycles ==> calculated cycles = dft*sinc2*sinc3 / adc Clk
                data_array[a, 16] = data_array[a, DFT_VALUE] / data_array[a, CYCLES];                                   //calculate samples per cycle
                data_array[a, SAMP_FREQ] = 800000 / data_array[a, SINC2_VALUE] * data_array[a, SINC3_VALUE];                          //calculate sampling freq
            }                                                            //assign column 11, 12, 13 to cycles, samp/cycle, samp freq

            //create an array of 24 excel worksheets 
            Excel.Worksheet[] ws = new Excel.Worksheet[25];

            //create a missing value object
            object misValue = System.Reflection.Missing.Value;                                           //for graphing

            //instantiate an excel application object
            Excel.Application my_excel_chart = new Excel.Application();
            my_excel_chart.Visible = true;                                                               //make the object visible

            //instantiate a workbook object
            Excel.Workbook wb = my_excel_chart.Workbooks.Add(misValue);



            //instantiate 24 worksheet object
            ws[1] = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            for (int j = 2; j < num_of_sheets; j++)
            {
                ws[j] = (Excel.Worksheet)wb.Worksheets.Add();
            }

            //instantiate another worksheet object to hold ALL the cared about data on one sheet
            Excel.Worksheet ws26 = (Excel.Worksheet)wb.Worksheets.Add();                                    //Create a worksheet for all of the graphs on one sheet

            //FOR all 25 Worksheets in Workbook1
            for (int k = 1; k < num_of_sheets; k++)
            {

                //Parse the Array for each sweep
                {
                    if (k == 1)
                    {
                        for (int x = 0; x < xl_length; x++)                                           //data_array1 is the first sweep
                        {
                            for (int y = 0; y < xl_width; y++)
                            {
                                data_array1[x, y] = data_array[x, y];
                                Data_Array[1] = data_array[x, y];
                            }
                        }
                    }
                    else if (k > 1)
                    {
                        //copy parts of the data array to another array to be printed on different sheets
                        for (int x = (k - 1) * xl_length; x < (k) * xl_length; x++)                                           //copy 57 rows to an array
                        {
                            for (int y = 0; y < xl_width; y++)
                            {
                                data_array1[x - ((k - 1) * xl_length), y] = data_array[x - (k - 1), y];
                            }
                        }
                    }
                }


                //instantiate a range 
                Excel.Range rng = ws[k].Cells.get_Resize(data_array1.GetLength(0), data_array1.GetLength(1));       //get the range
                rng.Value = data_array1;                                                                         //set the range to my data array 

                //Add Borders
                {
                    Excel.Range border_rng5 = ws[k].Range["L1:F57"];       //Phase
                    Excel.Range border_rng6 = ws[k].Range["F1:L1"];       //All heading
                    Excel.Borders border1 = border_rng5.Borders;
                    border1 = border_rng5.Borders;
                    border1.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border1.Weight = Excel.XlBorderWeight.xlThick;
                    border1.Color = Color.Blue;
                    Excel.Borders border2 = border_rng6.Borders;
                    border2.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border2.Weight = Excel.XlBorderWeight.xlThick;
                    border2.Color = Color.Red;
                }

                //make all sheets visable
                ws[k].Visible = Excel.XlSheetVisibility.xlSheetVisible;


                {
                    ws[k].Cells[1, 1] = "DFT_Table";
                    ws[k].Cells[1, 2] = "Sinc2_Table";
                    ws[k].Cells[1, 3] = "Sinc3_Table";
                    ws[k].Cells[1, 4] = "DFT Src";
                    ws[k].Cells[1, 5] = "PwrMod_Table";
                    ws[k].Cells[1, 6] = "Voltage Magnitude";
                    ws[k].Cells[1, 7] = "Voltage Phase";
                    ws[k].Cells[1, 8] = "Current Magnitude";
                    ws[k].Cells[1, 9] = "Current Phase";
                    ws[k].Cells[1, 10] = "Frequency";
                    ws[k].Cells[1, 11] = "Magnitude";
                    ws[k].Cells[1, 12] = "Phase";
                    ws[k].Cells[1, 13] = "DFT Num";
                    ws[k].Cells[1, 14] = "Sinc2";
                    ws[k].Cells[1, 15] = "Sinc3";
                    ws[k].Cells[1, 16] = "Cycles";
                    ws[k].Cells[1, 17] = "Samples/Cycle";
                    ws[k].Cells[1, 18] = "Sampling Freq";

                }       //set the excel sheet labels

                {
                    //Plot the magnitude////////////////////plot it on its own sheet. and also print it on the sheet with all of the plots/////////////
                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)ws[k].ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(700, 80, 600, 400);
                    Excel.Range oRng = ws[k].get_Range("I1", "I60");
                    Excel.Chart ct = myChart.Chart;
                    var missing = System.Type.Missing;
                    ct.ChartWizard(oRng, Excel.XlChartType.xl3DLine, Excel.XlScaleType.xlScaleLogarithmic);
                    Excel.Series oSeries = (Excel.Series)ct.SeriesCollection(1);
                    oSeries.XValues = ws[k].get_Range("H2", "H60");
                    //Microsoft.Office.Interop.Excel.XlScaleType ScaleType = Excel.XlScaleType.xlScaleLogarithmic;

                    ct.HasTitle = true;
                    ct.ChartTitle.Text = "Magnitude vs Frequency";
                    var yAxis = (Excel.Axis)ct.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    ct.Refresh();

                    //Plot the Phase///////////////////
                    Excel.ChartObjects xlCharts2 = (Excel.ChartObjects)ws[k].ChartObjects(Type.Missing);
                    Excel.ChartObject myChart2 = (Excel.ChartObject)xlCharts.Add(700, 480, 600, 400);
                    Excel.Range oRng2 = ws[k].get_Range("J1", "J60"); ;
                    Excel.Chart ct2 = myChart2.Chart;
                    ct2.ChartWizard(oRng2, Excel.XlChartType.xl3DLine, Excel.XlScaleType.xlScaleLogarithmic);
                    Excel.Series oSeries2 = (Excel.Series)ct2.SeriesCollection(1);
                    oSeries2.XValues = ws[k].get_Range("H2", "H60");
                    ct2.HasTitle = true;
                    ct2.ChartTitle.Text = "Phase vs Frequency";
                    ct2.Refresh();
                }       //plot the magnitude and phase for the first sheet


                get_electrode_labels(k);

                //Plot the magnitude////////////////////plot it on its own sheet. and also print it on the sheet with all of the plots/////////////
                Excel.ChartObjects xlCharts26 = (Excel.ChartObjects)ws26.ChartObjects(Type.Missing);
                Excel.ChartObject myChart26 = (Excel.ChartObject)xlCharts26.Add((0 + k - 1) * 290, 0, 350, 200);
                Excel.Range oRng26 = ws[k].get_Range("I1", "I60");
                Excel.Chart ct26 = myChart26.Chart;
                ct26.ChartWizard(oRng26, Excel.XlChartType.xl3DLine, Excel.XlScaleType.xlScaleLogarithmic);
                Excel.Series oSeries26 = (Excel.Series)ct26.SeriesCollection(1);
                oSeries26.XValues = ws[k].get_Range("H2", "H60");
                //Microsoft.Office.Interop.Excel.XlScaleType ScaleType = Excel.XlScaleType.xlScaleLogarithmic;
                ct26.HasTitle = true;
                ct26.ChartTitle.Text = "Magnitude vs Frequency " + electrode_label;
                var yAxis26 = (Excel.Axis)ct26.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                ct26.Refresh();

                //Now plot the Phase
                Excel.ChartObjects xlCharts2_26 = (Excel.ChartObjects)ws26.ChartObjects(Type.Missing);
                Excel.ChartObject myChart2_26 = (Excel.ChartObject)xlCharts2_26.Add((0 + k - 1) * 290, 200, 350, 200);
                Excel.Range oRng2_26 = ws[k].get_Range("J1", "J60");
                Excel.Chart ct2_26 = myChart2_26.Chart;
                ct2_26.ChartWizard(oRng2_26, Excel.XlChartType.xl3DLine, Excel.XlScaleType.xlScaleLogarithmic);
                Excel.Series oSeries2_26 = (Excel.Series)ct2_26.SeriesCollection(1);
                oSeries2_26.XValues = ws[k].get_Range("H2", "H60");
                ct2_26.HasTitle = true;
                ct2_26.ChartTitle.Text = "Phase vs Frequency" + electrode_label;
                ct2_26.Refresh();

                //instantiate a range 
                // Excel.Range rng = ws[k].Cells.get_Resize(xl_length,5);       //get the range
                // rng.Value = ws[k].Range["J1:F57"];


              //for (int l = 1; l < xl_length + 2; l++)
              //{
              //    ws26.Cells[30 + l, ((k - 1) * 6 + 1)] = ws[k].Cells[2 + l - 1, 6];   //set row 31 to (31+57) respectivley to each worksheets row 2 to 56     //do it for columns 6-10
              //    ws26.Cells[30 + l, ((k - 1) * 6 + 2)] = ws[k].Cells[2 + l - 1, 7];   //set row 31 to (31+57) respectivley to each worksheets row 2 to 56     //set column 0-5 on ws26  to  column 6 to 10 for each sheet
              //    ws26.Cells[30 + l, ((k - 1) * 6 + 3)] = ws[k].Cells[2 + l - 1, 8];   //set row 31 to (31+57) respectivley to each worksheets row 2 to 56     //set column 0-5 on ws26  to  column 6 to 10 for each sheet
              //    ws26.Cells[30 + l, ((k - 1) * 6 + 4)] = ws[k].Cells[2 + l - 1, 9];   //set row 31 to (31+57) respectivley to each worksheets row 2 to 56     //set column 0-5 on ws26  to  column 6 to 10 for each sheet
              //    ws26.Cells[30 + l, ((k - 1) * 6 + 5)] = ws[k].Cells[2 + l - 1, 10];   //set row 31 to (31+57) respectivley to each worksheets row 2 to 56     //set column 0-5 on ws26  to  column 6 to 10 for each sheet
              //
              //    
              //
              //
              // 
              //
              //
              //}

                //Copy data from each sheet to a main sheet
                //string columnLetter = ColumnIndexToColumnLetter(100); // returns "CV"
                string columnLetter1 = ColumnIndexToColumnLetter((k - 1) * 11 + 1); // returns the column string value
                string columnLetter2 = ColumnIndexToColumnLetter((k - 1) * 11 + 7); // returns the column string value
                Excel.Range from = ws[k].get_Range("F2", "L56");
                Excel.Range to = ws26.get_Range(columnLetter1 + "31", columnLetter1 + "85");
                from.Copy(to);

                ws26.Cells[29, (k - 1) * 11 + 1] = electrode_label;
                ws26.Cells[30, (k - 1) * 11 + 1] =  "|V|";
                ws26.Cells[30, (k - 1) * 11 + 2] =  "V Phase";
                ws26.Cells[30, (k - 1) * 11 + 3] =  "|I|";
                ws26.Cells[30, (k - 1) * 11 + 4] =  "I Phase";
                ws26.Cells[30, (k - 1) * 11 + 5] =  "Freq";
                ws26.Cells[30, (k - 1) * 11 + 6] =  "|Z|";
                ws26.Cells[30, (k - 1) * 11 + 7] =  "Z Phase";
                

            }


           //Excel.Range border_rng1_26 = ws26.Range["A31:ET57"];       //all
           //Excel.Range border_rng2_26 = ws26.Range["A30:ET30"];       //All heading
           //Excel.Borders border1_26 = border_rng1_26.Borders;
           //border1_26.LineStyle = Excel.XlLineStyle.xlContinuous;
           //border1_26.Weight = Excel.XlBorderWeight.xlThick;
           //border1_26.Color = Color.Red;
           //Excel.Borders border2_26 = border_rng2_26.Borders;
           //border2_26.LineStyle = Excel.XlLineStyle.xlContinuous;
           //border2_26.Weight = Excel.XlBorderWeight.xlThick;
           //border2_26.Color = Color.Blue;

           
        
        }
        private void button3_Click_1(object sender, EventArgs e)                               //Send to EXCEL 
        {
            //  using (Form3 frm = new Form3())
            //  {
            //      frm.Show();
            //  }

            MessageBox.Show("Allow extra time wihen the Excel Sheet loads to allow Excel to finish buffering");
           // using (Form3 frm3 = new Form3(SaveData))
           //   {
           //       frm3.ShowDialog(this);
           //   }

            //data_array[10000, 11];
            //C0 = dftnum table
            //C1 = sinc2 table
            //C2 = sinc3 table
            //C3 = dft src
            //C4 = pwr mod
            //C5 = Volt Mag
            //C6 = Volt phase
            //C7 = Curr Mag
            //C8 = Curr phase
            //C9 = freq result
            //C10 = Mag result
            //C11 = phase result
            //C12 = translated dft value
            //C13 = translated sinc2 value
            //C14= translated sinc3 value
            //C15= calculated cycles = dft*sinc2*sinc3 / adc Clk
            //C16= samples/cycle = dft / cycles
            //c17= sampling freq = adc clk / (sinc2*sinc3)

           // MessageBox.Show("Allow extra time wihen the Excel Sheet loads to allow Excel to finish buffering");
            SaveData();

            

          //  using (Form3 frm = new Form3())
          //  {
          //      frm.Hide();
          //  }
        }
        static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
        private void label13_Click(object sender, EventArgs e)                                 //Nothing
        {

        }
        private void label5_Click(object sender, EventArgs e)                                   //Nothing
        {

        }

        private void button1_Click_2(object sender, EventArgs e)                                //Button to send control to MCU to do a partial sweep
        {
            char[] sendToChip = new char[1];

            try
            {
                switch (comboBox8.Text)
                {
                    case "1 - 200k Hz":
                        {
                            sendToChip[0] = '!';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    case "1 - 10 Hz":
                        {
                            sendToChip[0] = '@';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    case "10 - 100 Hz":
                        {
                            sendToChip[0] = '#';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    case "100 - 1k Hz":
                        {
                            sendToChip[0] = '(';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    case "1k - 10k Hz":
                        {
                            sendToChip[0] = '%';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    case "10k - 100k Hz":
                        {
                            sendToChip[0] = '^';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    case "100k - 200k Hz":
                        {
                            sendToChip[0] = '&';
                            _setting._serial.Write(sendToChip, 0, 1);
                            quickPause();
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("A Sweep Was Not Sent To MCU");
                            break;
                        }
                }
            }
            catch
            {
                MessageBox.Show("Please select a sweep parameter");
            }
        }


        //////////SWICHES COMMANDS to be interpertated by AFE///////////////////
        ///DE0 = 5
        ///AFE1 = 6
        ///CE0 = 7
        ///AIN0 = 4
        ///AIN1 = 1
        ///AIN2 = 2 
        ///AIN3 = 3 
        //////////////////////////////////////////////////////////////////////
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)                     //determine switch 1
        {
            switch (comboBox1.Text)
            {
                case "CE0":
                    {
                        switch1 = '7';
                        break;
                    }
                case "AIN1":
                    {
                        switch1 = '1';
                        break;
                    }
                case "AIN2":
                    {
                        switch1 = '2';
                        break;

                    }
                case "AIN3":
                    {
                        switch1 = '3';
                        break;
                    }
                case "AFE1":
                    {
                        switch1 = '6';
                        break;
                    }
                default:
                    {
                        MessageBox.Show("A Switch Config was not matched");
                        break;
                    }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)                     //determine switch 2
        {
            switch (comboBox2.Text)
            {
                case "CE0":
                    {
                        switch2 = '7';
                        break;
                    }
                case "AIN1":
                    {
                        switch2 = '1';
                        break;
                    }
                case "AIN2":
                    {
                        switch2 = '2';
                        break;

                    }
                case "AIN3":
                    {
                        switch2 = '3';
                        break;
                    }
                case "DE0":
                    {
                        switch2 = '5';
                        break;
                    }
                default:
                    {
                        MessageBox.Show("A Switch Config was not matched");
                        break;
                    }
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)                     //nothing
        {


        }

        private void label15_Click_1(object sender, EventArgs e)                                    //nothing
        {

        }

        private void label22_Click(object sender, EventArgs e)                                      //nothing
        {

        }

        private void V_button_CheckedChanged(object sender, EventArgs e)                            //Make the vltage lables visable
        {
            V.BringToFront();
            V_label1.Visible = true;
            V_label2.Visible = true;
            I_label1.Visible = false;
            I_label2.Visible = false;
            bring_labels_to_front();

            switch1 = '6';                  //set the D switch to "6" which indicates ***AFE1*** for the MCU

        }

        private void I_button_CheckedChanged(object sender, EventArgs e)                            //Make Current lables visable
        {
            I.BringToFront();
            V_label1.Visible = false;
            V_label2.Visible = false;
            I_label1.Visible = true;
            I_label2.Visible = true;
            bring_labels_to_front();

            switch1 = '7';                  //set the D switch to "6" which indicates ***CE0*** for the MCU

        }

        private void button4_Click(object sender, EventArgs e)                                      //display the correct image to reflect the electrodes active
        {
            if (V_button.Checked == true)
            {
                switch (pos_ref.Text + neg_ref.Text)
                {
                    case "10":
                    case "01":
                        {
                            V01.BringToFront();
                            break;
                        }
                    case "20":
                    case "02":
                        {
                            V02.BringToFront();
                            break;
                        }
                    case "30":
                    case "03":
                        {
                            V03.BringToFront();
                            break;
                        }
                    case "21":
                    case "12":
                        {
                            V12.BringToFront();
                            break;
                        }
                    case "31":
                    case "13":
                        {
                            V31.BringToFront();
                            break;
                        }
                    case "32":
                    case "23":
                        {
                            V32.BringToFront();
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("A Measuremnt Reference Config was not matched");
                            break;
                        }

                }

            }
            if (I_button.Checked == true)
            {
                switch (pos_ref.Text + neg_ref.Text)
                {
                    case "10":
                    case "01":
                        {
                            I01.BringToFront();
                            break;
                        }
                    case "20":
                    case "02":
                        {
                            I02.BringToFront();
                            break;
                        }
                    case "30":
                    case "03":
                        {
                            I03.BringToFront();
                            break;
                        }
                    case "21":
                    case "12":
                        {
                            I12.BringToFront();
                            break;
                        }
                    case "31":
                    case "13":
                        {
                            I13.BringToFront();
                            break;
                        }
                    case "32":
                    case "23":
                        {
                            I23.BringToFront();
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("A Measuremnt Reference Config was not matched");
                            break;
                        }
                }
            }
            bring_labels_to_front();

            char[] sendToChip = new char[1];
            //Send switch config control key to MCU
            {
                sendToChip[0] = 'X';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send first Switch1 config to MCU
            {
                sendToChip[0] = switch1;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            //send  Switch2 config to MCU
            {
                sendToChip[0] = switch2;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            //send  Switch2 config to MCU
            {
                sendToChip[0] = switch3;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            //send  Switch2 config to MCU
            {
                sendToChip[0] = switch4;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send control key to end recording
            {
                sendToChip[0] = 'X';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
        }

        private void button2_Click_2(object sender, EventArgs e)                                    //send the switch data to the MCU
        {

            char[] sendToChip = new char[1];
            //Send Frequency control key to MCU
            {
                sendToChip[0] = 'X';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send first Switch1 config to MCU
            {
                sendToChip[0] = switch1;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            //send  Switch2 config to MCU
            {
                sendToChip[0] = switch2;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            //send  Switch2 config to MCU
            {
                sendToChip[0] = switch3;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            //send  Switch2 config to MCU
            {
                sendToChip[0] = switch4;
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send comma to seperate the next config
            {
                sendToChip[0] = ',';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send control key to end recording
            {
                sendToChip[0] = 'X';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
        }

        private void pos_ref_SelectedIndexChanged(object sender, EventArgs e)                       //Set the + ref switch (P)          //the pos_ref switch is **switch P** for the MCU  ==> switch 2
        {
            switch (pos_ref.Text)
            {
                case "0":                       //electrode 0 = swDE0 for the pos
                    {
                        switch2 = '5';          //the code 5 means DEO to the AFE
                        break;
                    }
                case "1":                       //electrode 1 = swDAIN1 for the pos
                    {
                        switch2 = '1';          //the code 1 means AIN1 to the AFE
                        break;
                    }
                case "2":                       //electrode 2 = swDAIN2 for the pos
                    {
                        switch2 = '2';          //the code 2 means AIN2 to the AFE
                        break;

                    }
                case "3":                       //electrode 3 = swDAIN3 for the pos
                    {
                        switch2 = '3';          //the code 3 means AIN3 to the AFE
                        break;
                    }
                default:
                    {
                        MessageBox.Show("A Switch Config was not matched");
                        break;
                    }
            }

        }

        private void neg_ref_SelectedIndexChanged(object sender, EventArgs e)                       //the neg_ref switch is **switch N** for the MCU  ==> switch 3
        {
            switch (neg_ref.Text)
            {
                case "0":                       //electrode 0 = swAIN0 for the pos
                    {
                        switch3 = '4';          //set the neg ref switch - the code 4 means AIN0 to the AFE
                        switch4 = '4';          //set the T switch - the code 4 means AIN0 to the AFE
                        break;
                    }
                case "1":                       //electrode 1 = swDAIN1 for the pos
                    {
                        switch3 = '1';          //the code 1 means AIN1 to the AFE
                        switch4 = '1';          //set the T switch - the code 4 means AIN0 to the AFE

                        break;
                    }
                case "2":                       //electrode 2 = swDAIN2 for the pos
                    {
                        switch3 = '2';          //the code 2 means AIN2 to the AFE
                        switch4 = '2';          //set the T switch - the code 4 means AIN0 to the AFE
                        break;

                    }
                case "3":                       //electrode 3 = swDAIN3 for the pos
                    {
                        switch3 = '3';          //the code 3 means AIN3 to the AFE
                        switch4 = '3';          //set the T switch - the code 4 means AIN0 to the AFE
                        break;
                    }
                default:
                    {
                        MessageBox.Show("A Switch Config was not matched");
                        break;
                    }
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)                     //Set switch 3
        {
            {
                switch (comboBox6.Text)
                {
                    case "AIN0":
                        {
                            switch3 = '4';
                            break;
                        }
                    case "AIN1":
                        {
                            switch3 = '1';
                            break;
                        }
                    case "AIN2":
                        {
                            switch3 = '2';
                            break;

                        }
                    case "AIN3":
                        {
                            switch3 = '3';
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("A Switch Config was not matched");
                            break;
                        }
                }
            }
        }

        private void richTextBox1_TextChanged_1(object sender, EventArgs e)                         //Nothing
        {

        }

        private void button3_Click(object sender, EventArgs e)                                      //Button to send control to MCU for entire 24 sweep
        {
            char[] sendToChip = new char[1];

            {
                sendToChip[0] = 'E';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
        }

        private void sweep_number_label_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void V_label2_Click(object sender, EventArgs e)
        {

        }

        private void I_label2_Click(object sender, EventArgs e)
        {

        }

        private void I_label1_Click(object sender, EventArgs e)
        {

        }

        private void V_label1_Click(object sender, EventArgs e)
        {

        }

        private void E1label_Click(object sender, EventArgs e)
        {

        }

        private void E2label_Click(object sender, EventArgs e)
        {

        }

        private void E3label_Click(object sender, EventArgs e)
        {

        }

        private void E0label_Click(object sender, EventArgs e)
        {

        }

        private void STOCK_img_Click(object sender, EventArgs e)
        {

        }

        private void V_Click(object sender, EventArgs e)
        {

        }

        private void I_Click(object sender, EventArgs e)
        {

        }

        private void V01_Click(object sender, EventArgs e)
        {

        }

        private void V02_Click(object sender, EventArgs e)
        {

        }

        private void V03_Click(object sender, EventArgs e)
        {

        }

        private void V12_Click(object sender, EventArgs e)
        {

        }

        private void V31_Click(object sender, EventArgs e)
        {

        }

        private void V32_Click(object sender, EventArgs e)
        {

        }

        private void I01_Click(object sender, EventArgs e)
        {

        }

        private void I02_Click(object sender, EventArgs e)
        {

        }

        private void I03_Click(object sender, EventArgs e)
        {

        }

        private void I12_Click(object sender, EventArgs e)
        {

        }

        private void I13_Click(object sender, EventArgs e)
        {

        }

        private void I23_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void final_pk2pk_textbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void dac_gain_button1_CheckedChanged(object sender, EventArgs e)
        {
            if (dac_gain_button1.Checked)
            {
                dacgain = 0.2;
                dacgaincode = '2';
            }
            if(dac_gain_button2.Checked)
            {
                dacgain = 1;
                dacgaincode = '1';
            }

            finalpk2pk = Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain;
            final_pk2pk_textbox.Text = Convert.ToString(finalpk2pk);
            DrawSine();
        }

        private void ExBuffGain_button1_CheckedChanged(object sender, EventArgs e)
        {
            if (ExBuffGain_button1.Checked)
            {
                exbuffgain = 0.25;
                exbuffgaincode ='1';              //this will be sent to the chip so it knows what to set the ex buff gain to
            }
            if (ExBuffGain_button2.Checked)
            {
                exbuffgain = 2;
                exbuffgaincode = '2';
            }

            finalpk2pk = Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain;
            final_pk2pk_textbox.Text = Convert.ToString(finalpk2pk);
            DrawSine();
        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void helpMenuToolStripMenuItem_Click(object sender, EventArgs e)        //clicked on the help menue
        {
            _helpmenue.Show();
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void ExBuffGain_button2_CheckedChanged(object sender, EventArgs e)
        {
            if (ExBuffGain_button1.Checked)
            {
                exbuffgain = 0.25;
                exbuffgaincode = '1';              //this will be sent to the chip so it knows what to set the ex buff gain to
            }
            if (ExBuffGain_button2.Checked)
            {
                exbuffgain = 2;
                exbuffgaincode = '2';
            }

            finalpk2pk = Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain;
            final_pk2pk_textbox.Text = Convert.ToString(finalpk2pk);
            DrawSine();
        }

        private void offset_scroll_Scroll(object sender, ScrollEventArgs e)
        {
            excitationOffset = offset_scroll.Value;
            richTextBox2.Text = Convert.ToString( offset_scroll.Value);
            DrawSine();

        }

        private void dac_gain_button2_CheckedChanged(object sender, EventArgs e)
        {
            if (dac_gain_button1.Checked)
            {
                dacgain = 0.2;
                dacgaincode = '2';
            }
            if (dac_gain_button2.Checked)
            {
                dacgain = 1;
                dacgaincode = '1';
            }

            finalpk2pk = Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain;
            final_pk2pk_textbox.Text = Convert.ToString(finalpk2pk);
            DrawSine();
        }

        private void disconnectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                _setting._serial.Close();
                _setting._serial2.Close();
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Bluetooth Connection";
                toolStripStatusLabel1.BackColor = Color.Black;
            }
            catch
            {

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            char[] sendToChip = new char[1];
            frequency = richTextBox4.Text;

            //error handling, the pk-pk for the dac has to be: ((dac pk-pk) * (dac buffer) <= 1800mv
            if (Convert.ToDouble(pk2pk_textbox.Text) * dacgain   > 1799)
            {
                MessageBox.Show("The DAC pk-pk * DAC Gain must be less than 1800mV");
            }

            // pk-pk for the excitation buffer has to be: ((dac pk-pk) * (dac buffer) * (ex. buff gain) <= 2200mv
            else if (Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain > 2199)
            {
                MessageBox.Show("The DAC pk-pk * DAC Gain * Excitation Buffer Gain must be less than 2200mV");
            }

            else
            {
                //No errors, send info to chip
                //Send Frequency control key to MCU
                {
                    sendToChip[0] = 'P';                            //send the characteer to start recording state for pk-pk and the gains
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }
                {
                    sendToChip[0] = dacgaincode;                            //send the characteer respective to the dac gain
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }
                {
                    sendToChip[0] = ',';                            //send the seperating
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }
                {
                    sendToChip[0] = exbuffgaincode;                            //send the characteer respective to the dac gain
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }
                {
                    sendToChip[0] = ',';                            //send the seperating
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }
                
                //send the dac pk-pk to MCU
                for (int i = 0; i < pk2pk_textbox.TextLength; i++)
                {
                    sendToChip[0] = Convert.ToChar(pk2pk_textbox.Text[i]);
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }

                //send stop recording control key to end recording
                {
                    sendToChip[0] = 'D';
                    _setting._serial.Write(sendToChip, 0, 1);
                    quickPause();
                }
            }
            
        }

        private void pk2pk_textbox_TextChanged(object sender, EventArgs e)
        {

            if (pk2pk_textbox.Text == "")
            { }
            else
            {
                finalpk2pk = Convert.ToDouble(pk2pk_textbox.Text) * dacgain * exbuffgain;       
                final_pk2pk_textbox.Text = Convert.ToString(finalpk2pk);
            }
            DrawSine();
           
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)                     //set switch 4 
        {
            {
                switch (comboBox7.Text)
                {
                    case "AIN0":
                        {
                            switch4 = '4';
                            break;
                        }
                    case "AIN1":
                        {
                            switch4 = '1';
                            break;
                        }
                    case "AIN2":
                        {
                            switch4 = '2';
                            break;

                        }
                    case "AIN3":
                        {
                            switch4 = '3';
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("A Switch Config was not matched");
                            break;
                        }
                }
            }
        }


        /// <summary>
        /// This is a function that takes one sample at the frequency that is specified in textbox4
        /// It sends the data to serial port 1
        /// </summary>
        /// THIS sends over the frequency, sinc2, sinc3, and DFTnum parameters to the MCU
        /// 
        public void TakeOneSample()
        {
            //Initialize local variables

            char[] sendToChip = new char[1];
            frequency = richTextBox4.Text;
            string sinc2_value;
            string sinc3_value;
            string dft_value;


            //Send Frequency control key to MCU
            {
                sendToChip[0] = 'F';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send the frequency to MCU
            for (int i = 0; i < richTextBox4.TextLength; i++)
            {
                sendToChip[0] = Convert.ToChar(richTextBox4.Text[i]);
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send freq control key to end recording
            {
                sendToChip[0] = 'F';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }


            /////////////Start sending the SINC3 /////////////////

            ///Send the Sinc3 command key to MCU to sart recording the Sinc3 value
            {
                sendToChip[0] = 'T';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            switch (Convert.ToInt32(comboBox4.Text))
            {
                case 5:
                    {
                        sinc3_value = "2";
                        break;
                    }
                case 4:
                    {
                        sinc3_value = "1";
                        break;
                    }
                case 2:
                    {
                        sinc3_value = "0";
                        break;
                    }
                default:
                    {
                        sinc3_value = "0";
                        break;
                    }
            }

            //Need to send the value to the MCU now
            for (int i = 0; i < sinc3_value.Length; i++)
            {
                sendToChip[0] = Convert.ToChar(sinc3_value[i]);
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send command saying that we have recorded all the parameters data
            {
                sendToChip[0] = 'T';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            ////Start sendinng the SINC2 parameter///////////////////

            //send the command to start recording the sinc2 parameter
            {
                sendToChip[0] = 'S';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            switch (Convert.ToInt32(comboBox5.Text))
            {
                case 22:
                    {
                        sinc2_value = "1";
                        break;
                    }
                case 44:
                    {
                        sinc2_value = "2";
                        break;
                    }
                case 89:
                    {
                        sinc2_value = "3";
                        break;
                    }
                case 178:
                    {
                        sinc2_value = "4";
                        break;
                    }
                case 267:
                    {
                        sinc2_value = "5";
                        break;
                    }
                case 533:
                    {
                        sinc2_value = "6";
                        break;
                    }
                case 640:
                    {
                        sinc2_value = "7";
                        break;
                    }
                case 667:
                    {
                        sinc2_value = "8";
                        break;
                    }
                case 800:
                    {
                        sinc2_value = "9";
                        break;
                    }
                case 889:
                    {
                        sinc2_value = "10";
                        break;
                    }
                case 1067:
                    {
                        sinc2_value = "11";
                        break;
                    }
                case 1333:
                    {
                        sinc2_value = "12";
                        break;
                    }
                default:
                    {
                        sinc2_value = "0";
                        break;
                    }
            }

            //send the sinc2 value to the MCU
            for (int i = 0; i < sinc2_value.Length; i++)
            {
                sendToChip[0] = Convert.ToChar(sinc2_value[i]);
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send command saying that we have recorded all the parameters data
            {
                sendToChip[0] = 'S';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            /////Start recoding the DFT NUM///////////////////////////

            //Send control key to start recording the DFT num
            {
                sendToChip[0] = 'D';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
            switch (Convert.ToInt32(comboBox3.Text))
            {

                case 4:
                    {
                        dft_value = "0";
                        break;
                    }
                case 8:
                    {
                        dft_value = "1";
                        break;
                    }
                case 16:
                    {
                        dft_value = "2";
                        break;
                    }
                case 32:
                    {
                        dft_value = "3";
                        break;
                    }
                case 64:
                    {
                        dft_value = "4";
                        break;
                    }
                case 128:
                    {
                        dft_value = "5";
                        break;
                    }
                case 256:
                    {
                        dft_value = "6";
                        break;
                    }
                case 512:
                    {
                        dft_value = "7";
                        break;
                    }
                case 1024:
                    {
                        dft_value = "8";
                        break;
                    }
                case 2048:
                    {
                        dft_value = "9";
                        break;
                    }
                case 4096:
                    {
                        dft_value = "10";
                        break;
                    }
                case 8192:
                    {
                        dft_value = "11";
                        break;
                    }
                case 16384:
                    {
                        dft_value = "12";
                        break;
                    }
                default:
                    {
                        dft_value = "0";
                        break;
                    }
            }

            //send the DFT value over to the MCU 
            for (int i = 0; i < dft_value.Length; i++)
            {
                sendToChip[0] = Convert.ToChar(dft_value[i]);
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            //send D command saying that we have recorded all of the parameters data
            {
                sendToChip[0] = 'D';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }

            ////////Sending all parameters is complete/////////////////////

            //Now sendd the H character to notify ALL parameters have been recorded
            {
                quickPause();
                sendToChip[0] = 'H';
                _setting._serial.Write(sendToChip, 0, 1);
                quickPause();
            }
        }


        /// <summary>
        /// This is a function that takes one sample at the frequency that is specified from the serial Port 2
        /// It sends the data to serial port 1 (3029MCU)
        /// </summary>
        


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public void bring_labels_to_front()                                                            //bring the lables to the front
        {
            V_label1.BringToFront();
            V_label2.BringToFront();
            I_label1.BringToFront();
            I_label2.BringToFront();
            E0label.BringToFront();
            E1label.BringToFront();
            E2label.BringToFront();
            E3label.BringToFront();
        }
        public string electrode_label;

        public void get_electrode_labels(int k)
        {
            switch (k)
            {
                case 1:
                    electrode_label = "A0B1";
                    break;
                case 2:
                    electrode_label = "A0B2";
                    break;
                case 3:
                    electrode_label = "A0B3";
                    break;               
                case 4:                  
                    electrode_label = "A1B0";
                    break;               
                case 5:                  
                    electrode_label = "A1B2";
                    break;               
                case 6:                  
                    electrode_label = "A1B3";
                    break;               
                case 7:                  
                    electrode_label = "A2B0";
                    break;               
                case 8:                  
                    electrode_label = "A2B1";
                    break;               
                case 9:                  
                    electrode_label = "A2B3";
                    break;               
                case 10:                 
                    electrode_label = "A3B0";
                    break;               
                case 11:                 
                    electrode_label = "A3B1";
                    break;               
                case 12:                 
                    electrode_label = "A3B2";
                    break;
            }

        }
    }
}
