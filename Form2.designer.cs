namespace UART_Senior_Design_Test
{
    partial class Bluetooth_Settings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.com_port = new System.Windows.Forms.ComboBox();
            this.baud_rate = new System.Windows.Forms.ComboBox();
            this.data = new System.Windows.Forms.ComboBox();
            this.parity = new System.Windows.Forms.ComboBox();
            this.stop_bits = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.connect = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.com_port2 = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // com_port
            // 
            this.com_port.FormattingEnabled = true;
            this.com_port.Location = new System.Drawing.Point(81, 65);
            this.com_port.Name = "com_port";
            this.com_port.Size = new System.Drawing.Size(196, 28);
            this.com_port.TabIndex = 0;
            this.com_port.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // baud_rate
            // 
            this.baud_rate.FormattingEnabled = true;
            this.baud_rate.Items.AddRange(new object[] {
            "115200",
            "9660"});
            this.baud_rate.Location = new System.Drawing.Point(183, 123);
            this.baud_rate.Name = "baud_rate";
            this.baud_rate.Size = new System.Drawing.Size(196, 28);
            this.baud_rate.TabIndex = 1;
            this.baud_rate.Text = "115200";
            this.baud_rate.SelectedIndexChanged += new System.EventHandler(this.baud_rate_SelectedIndexChanged);
            // 
            // data
            // 
            this.data.FormattingEnabled = true;
            this.data.Items.AddRange(new object[] {
            "8 bit",
            "7 bit"});
            this.data.Location = new System.Drawing.Point(183, 186);
            this.data.Name = "data";
            this.data.Size = new System.Drawing.Size(196, 28);
            this.data.TabIndex = 2;
            this.data.Text = "8 bit";
            // 
            // parity
            // 
            this.parity.FormattingEnabled = true;
            this.parity.Items.AddRange(new object[] {
            "none",
            "odd",
            "even",
            "mark",
            "space"});
            this.parity.Location = new System.Drawing.Point(183, 249);
            this.parity.Name = "parity";
            this.parity.Size = new System.Drawing.Size(196, 28);
            this.parity.TabIndex = 3;
            this.parity.Text = "none";
            // 
            // stop_bits
            // 
            this.stop_bits.FormattingEnabled = true;
            this.stop_bits.Items.AddRange(new object[] {
            "1 bit",
            "1.5 bit",
            "2 bit"});
            this.stop_bits.Location = new System.Drawing.Point(183, 311);
            this.stop_bits.Name = "stop_bits";
            this.stop_bits.Size = new System.Drawing.Size(196, 28);
            this.stop_bits.TabIndex = 4;
            this.stop_bits.Text = "1 bit";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(116, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(153, 25);
            this.label1.TabIndex = 5;
            this.label1.Text = "COM PORT 1";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(44, 123);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(122, 25);
            this.label2.TabIndex = 6;
            this.label2.Text = "Baud Rate";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(44, 186);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 25);
            this.label3.TabIndex = 7;
            this.label3.Text = "Data";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(44, 249);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 25);
            this.label4.TabIndex = 8;
            this.label4.Text = "Parity";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(44, 311);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 25);
            this.label5.TabIndex = 9;
            this.label5.Text = "Stop Bits";
            // 
            // connect
            // 
            this.connect.Location = new System.Drawing.Point(216, 378);
            this.connect.Name = "connect";
            this.connect.Size = new System.Drawing.Size(132, 49);
            this.connect.TabIndex = 10;
            this.connect.Text = "Connect";
            this.connect.UseVisualStyleBackColor = true;
            this.connect.Click += new System.EventHandler(this.Connect_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(326, 32);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(153, 25);
            this.label6.TabIndex = 11;
            this.label6.Text = "COM PORT 2";
            // 
            // com_port2
            // 
            this.com_port2.FormattingEnabled = true;
            this.com_port2.Location = new System.Drawing.Point(285, 65);
            this.com_port2.Name = "com_port2";
            this.com_port2.Size = new System.Drawing.Size(196, 28);
            this.com_port2.TabIndex = 12;
            this.com_port2.Text = "COM99";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(404, 140);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(156, 203);
            this.button1.TabIndex = 13;
            this.button1.Text = "Auto Connect \r\n2 COM Ports";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(404, 378);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(118, 74);
            this.button2.TabIndex = 14;
            this.button2.Text = "BLE";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Bluetooth_Settings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(576, 477);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.com_port2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.connect);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.stop_bits);
            this.Controls.Add(this.parity);
            this.Controls.Add(this.data);
            this.Controls.Add(this.baud_rate);
            this.Controls.Add(this.com_port);
            this.Name = "Bluetooth_Settings";
            this.Text = "Bluetooth_Settings";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox com_port;
        private System.Windows.Forms.ComboBox baud_rate;
        private System.Windows.Forms.ComboBox data;
        private System.Windows.Forms.ComboBox parity;
        private System.Windows.Forms.ComboBox stop_bits;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button connect;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox com_port2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}