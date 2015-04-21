namespace WordCommentExtractor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenWordFile = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Word Files|*.doc;*.docx";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Select one or more Word files";
            // 
            // btnOpenWordFile
            // 
            this.btnOpenWordFile.Location = new System.Drawing.Point(67, 161);
            this.btnOpenWordFile.Name = "btnOpenWordFile";
            this.btnOpenWordFile.Size = new System.Drawing.Size(103, 23);
            this.btnOpenWordFile.TabIndex = 1;
            this.btnOpenWordFile.Text = "Select Word File";
            this.btnOpenWordFile.UseVisualStyleBackColor = true;
            this.btnOpenWordFile.Click += new System.EventHandler(this.btnOpenWordFile_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(21, 9);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(202, 115);
            this.textBox1.TabIndex = 5;
            this.textBox1.Text = resources.GetString("textBox1.Text");
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(67, 225);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(103, 23);
            this.btnExit.TabIndex = 6;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(242, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Status:";
            // 
            // txtStatus
            // 
            this.txtStatus.Location = new System.Drawing.Point(245, 34);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ReadOnly = true;
            this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtStatus.Size = new System.Drawing.Size(656, 258);
            this.txtStatus.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(925, 322);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnOpenWordFile);
            this.Name = "Form1";
            this.Text = "Word Comment Extractor";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnOpenWordFile;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtStatus;
    }
}

