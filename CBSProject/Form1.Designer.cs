namespace CBSProject
{
    partial class Form1
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
            System.Windows.Forms.Button btnBrowse;
            this.btnDisplay = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.txtSheet = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.cboSheet = new System.Windows.Forms.ComboBox();
            this.txtFilename = new System.Windows.Forms.Label();
            this.YearMlbl = new System.Windows.Forms.Label();
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.btnDB = new System.Windows.Forms.Button();
            btnBrowse = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            btnBrowse.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnBrowse.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            btnBrowse.Location = new System.Drawing.Point(712, 217);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new System.Drawing.Size(75, 23);
            btnBrowse.TabIndex = 46;
            btnBrowse.Text = "Browse..";
            btnBrowse.UseVisualStyleBackColor = false;
            btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnDisplay
            // 
            this.btnDisplay.Location = new System.Drawing.Point(357, 266);
            this.btnDisplay.Name = "btnDisplay";
            this.btnDisplay.Size = new System.Drawing.Size(75, 23);
            this.btnDisplay.TabIndex = 0;
            this.btnDisplay.Text = "Display";
            this.btnDisplay.UseVisualStyleBackColor = true;
            this.btnDisplay.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(17, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(770, 189);
            this.dataGridView1.TabIndex = 1;
            // 
            // txtSheet
            // 
            this.txtSheet.AutoSize = true;
            this.txtSheet.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.txtSheet.Location = new System.Drawing.Point(14, 260);
            this.txtSheet.Name = "txtSheet";
            this.txtSheet.Size = new System.Drawing.Size(35, 13);
            this.txtSheet.TabIndex = 50;
            this.txtSheet.Text = "Sheet";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.textBox1.Location = new System.Drawing.Point(69, 217);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(637, 20);
            this.textBox1.TabIndex = 49;
            // 
            // cboSheet
            // 
            this.cboSheet.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.cboSheet.FormattingEnabled = true;
            this.cboSheet.Location = new System.Drawing.Point(69, 252);
            this.cboSheet.Name = "cboSheet";
            this.cboSheet.Size = new System.Drawing.Size(100, 21);
            this.cboSheet.TabIndex = 48;
            this.cboSheet.SelectedIndexChanged += new System.EventHandler(this.cboSheet_SelectedIndexChanged);
            // 
            // txtFilename
            // 
            this.txtFilename.AutoSize = true;
            this.txtFilename.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.txtFilename.Location = new System.Drawing.Point(14, 217);
            this.txtFilename.Name = "txtFilename";
            this.txtFilename.Size = new System.Drawing.Size(51, 13);
            this.txtFilename.TabIndex = 47;
            this.txtFilename.Text = "FileName";
            // 
            // YearMlbl
            // 
            this.YearMlbl.AutoSize = true;
            this.YearMlbl.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.YearMlbl.Location = new System.Drawing.Point(551, 266);
            this.YearMlbl.Name = "YearMlbl";
            this.YearMlbl.Size = new System.Drawing.Size(39, 13);
            this.YearMlbl.TabIndex = 52;
            this.YearMlbl.Text = "YYMM";
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Location = new System.Drawing.Point(596, 261);
            this.dateTimePicker.Name = "dateTimePicker";
            this.dateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker.TabIndex = 53;
            // 
            // btnDB
            // 
            this.btnDB.Location = new System.Drawing.Point(345, 323);
            this.btnDB.Name = "btnDB";
            this.btnDB.Size = new System.Drawing.Size(102, 23);
            this.btnDB.TabIndex = 54;
            this.btnDB.Text = "Add to Database";
            this.btnDB.UseVisualStyleBackColor = true;
            this.btnDB.Click += new System.EventHandler(this.btnDB_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnDB);
            this.Controls.Add(this.dateTimePicker);
            this.Controls.Add(this.YearMlbl);
            this.Controls.Add(this.txtSheet);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.cboSheet);
            this.Controls.Add(this.txtFilename);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnDisplay);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDisplay;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label txtSheet;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ComboBox cboSheet;
        private System.Windows.Forms.Label txtFilename;
        private System.Windows.Forms.Label YearMlbl;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.Button btnDB;
    }
}

