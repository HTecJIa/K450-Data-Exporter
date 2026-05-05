namespace K450DataExporter
{
    partial class Form2
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.listBoxInfFiles = new System.Windows.Forms.ListBox();
            this.btnOpen = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listBoxInfFiles);
            this.groupBox1.Location = new System.Drawing.Point(4, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(243, 140);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Selection .inf";
            // 
            // listBoxInfFiles
            // 
            this.listBoxInfFiles.FormattingEnabled = true;
            this.listBoxInfFiles.Location = new System.Drawing.Point(6, 22);
            this.listBoxInfFiles.Name = "listBoxInfFiles";
            this.listBoxInfFiles.Size = new System.Drawing.Size(229, 108);
            this.listBoxInfFiles.TabIndex = 0;
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(91, 158);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 23);
            this.btnOpen.TabIndex = 1;
            this.btnOpen.Text = "OK";
            this.btnOpen.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(172, 158);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AccessibleName = "";
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(251, 191);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select the Initialization File";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox listBoxInfFiles;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Button btnCancel;
    }
}