namespace WindowsFormsApplication3
{
    partial class mainForm
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
            this.goButton = new System.Windows.Forms.Button();
            this.tableNameComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dbFileTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dbBrowseButton = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // goButton
            // 
            this.goButton.Location = new System.Drawing.Point(346, 27);
            this.goButton.Name = "goButton";
            this.goButton.Size = new System.Drawing.Size(96, 23);
            this.goButton.TabIndex = 0;
            this.goButton.Text = "&Export";
            this.goButton.UseVisualStyleBackColor = true;
            this.goButton.Click += new System.EventHandler(this.goButton_Click);
            // 
            // tableNameComboBox
            // 
            this.tableNameComboBox.FormattingEnabled = true;
            this.tableNameComboBox.Location = new System.Drawing.Point(12, 79);
            this.tableNameComboBox.Name = "tableNameComboBox";
            this.tableNameComboBox.Size = new System.Drawing.Size(254, 21);
            this.tableNameComboBox.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Table name:";
            // 
            // dbFileTextBox
            // 
            this.dbFileTextBox.Location = new System.Drawing.Point(12, 27);
            this.dbFileTextBox.Name = "dbFileTextBox";
            this.dbFileTextBox.Size = new System.Drawing.Size(254, 20);
            this.dbFileTextBox.TabIndex = 3;
            this.dbFileTextBox.TextChanged += new System.EventHandler(this.dbFileTextBox_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Db file:";
            // 
            // dbBrowseButton
            // 
            this.dbBrowseButton.Location = new System.Drawing.Point(273, 27);
            this.dbBrowseButton.Name = "dbBrowseButton";
            this.dbBrowseButton.Size = new System.Drawing.Size(26, 20);
            this.dbBrowseButton.TabIndex = 5;
            this.dbBrowseButton.Text = "..";
            this.dbBrowseButton.UseVisualStyleBackColor = true;
            this.dbBrowseButton.Click += new System.EventHandler(this.dbBrowseButton_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(467, 222);
            this.Controls.Add(this.dbBrowseButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dbFileTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableNameComboBox);
            this.Controls.Add(this.goButton);
            this.Name = "mainForm";
            this.Text = "SqliteExporter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button goButton;
        private System.Windows.Forms.ComboBox tableNameComboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox dbFileTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button dbBrowseButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
    }
}

