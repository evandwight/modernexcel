namespace ModernExcel
{
    partial class MyUserControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.runButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.skipButton = new System.Windows.Forms.Button();
            this.indexText = new System.Windows.Forms.Label();
            this.applyButton = new System.Windows.Forms.Button();
            this.cellTextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.nameTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.messageBox = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // runButton
            // 
            this.runButton.Location = new System.Drawing.Point(32, 52);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(75, 34);
            this.runButton.TabIndex = 1;
            this.runButton.Text = "analyze";
            this.runButton.UseVisualStyleBackColor = true;
            this.runButton.Click += new System.EventHandler(this.runButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.skipButton);
            this.groupBox1.Controls.Add(this.indexText);
            this.groupBox1.Controls.Add(this.applyButton);
            this.groupBox1.Controls.Add(this.cellTextBox);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.nameTextBox);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(7, 182);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(239, 246);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Action";
            this.groupBox1.Visible = false;
            // 
            // skipButton
            // 
            this.skipButton.Location = new System.Drawing.Point(76, 122);
            this.skipButton.Name = "skipButton";
            this.skipButton.Size = new System.Drawing.Size(75, 30);
            this.skipButton.TabIndex = 4;
            this.skipButton.Text = "Skip";
            this.skipButton.UseVisualStyleBackColor = true;
            this.skipButton.Click += new System.EventHandler(this.skipButton_Click);
            // 
            // indexText
            // 
            this.indexText.AutoSize = true;
            this.indexText.Location = new System.Drawing.Point(78, 163);
            this.indexText.Name = "indexText";
            this.indexText.Size = new System.Drawing.Size(49, 20);
            this.indexText.TabIndex = 4;
            this.indexText.Text = "_ of _";
            // 
            // applyButton
            // 
            this.applyButton.Location = new System.Drawing.Point(157, 122);
            this.applyButton.Name = "applyButton";
            this.applyButton.Size = new System.Drawing.Size(76, 30);
            this.applyButton.TabIndex = 6;
            this.applyButton.Text = "Apply";
            this.applyButton.UseVisualStyleBackColor = true;
            this.applyButton.Click += new System.EventHandler(this.applyButton_Click);
            // 
            // cellTextBox
            // 
            this.cellTextBox.Location = new System.Drawing.Point(82, 80);
            this.cellTextBox.Name = "cellTextBox";
            this.cellTextBox.ReadOnly = true;
            this.cellTextBox.Size = new System.Drawing.Size(151, 26);
            this.cellTextBox.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 20);
            this.label3.TabIndex = 3;
            this.label3.Text = "Cell";
            // 
            // nameTextBox
            // 
            this.nameTextBox.Location = new System.Drawing.Point(82, 37);
            this.nameTextBox.Name = "nameTextBox";
            this.nameTextBox.Size = new System.Drawing.Size(151, 26);
            this.nameTextBox.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Name:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(245, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "Click analyze to find cells to name";
            // 
            // messageBox
            // 
            this.messageBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.messageBox.ForeColor = System.Drawing.Color.Red;
            this.messageBox.Location = new System.Drawing.Point(8, 90);
            this.messageBox.Multiline = true;
            this.messageBox.Name = "messageBox";
            this.messageBox.ReadOnly = true;
            this.messageBox.Size = new System.Drawing.Size(240, 86);
            this.messageBox.TabIndex = 5;
            this.messageBox.Visible = false;
            // 
            // MyUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.messageBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.runButton);
            this.Name = "MyUserControl";
            this.Size = new System.Drawing.Size(264, 450);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button applyButton;
        private System.Windows.Forms.TextBox cellTextBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox nameTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label indexText;
        private System.Windows.Forms.Button skipButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox messageBox;
    }
}
