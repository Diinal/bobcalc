namespace Bobcalc
{
    partial class Page1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Page1));
            this.MuteButton = new System.Windows.Forms.PictureBox();
            this.CloseButton = new System.Windows.Forms.PictureBox();
            this.Logo = new System.Windows.Forms.PictureBox();
            this.Top_Label = new System.Windows.Forms.Label();
            this.Start_Button = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem_Settings = new System.Windows.Forms.ToolStripMenuItem();
            this.путьДляExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.путьДляPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.MuteButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CloseButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Logo)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // MuteButton
            // 
            this.MuteButton.BackColor = System.Drawing.Color.AliceBlue;
            this.MuteButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.MuteButton.Image = global::Bobcalc.Properties.Resources.Свернуть4;
            this.MuteButton.Location = new System.Drawing.Point(890, 22);
            this.MuteButton.Name = "MuteButton";
            this.MuteButton.Size = new System.Drawing.Size(51, 50);
            this.MuteButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.MuteButton.TabIndex = 6;
            this.MuteButton.TabStop = false;
            this.MuteButton.Click += new System.EventHandler(this.MuteButton_Click_1);
            this.MuteButton.MouseLeave += new System.EventHandler(this.MuteButton_MouseLeave_1);
            this.MuteButton.MouseHover += new System.EventHandler(this.MuteButton_MouseHover_1);
            // 
            // CloseButton
            // 
            this.CloseButton.BackColor = System.Drawing.Color.AliceBlue;
            this.CloseButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CloseButton.Image = global::Bobcalc.Properties.Resources._1498934303_basics_22;
            this.CloseButton.Location = new System.Drawing.Point(960, 22);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(51, 50);
            this.CloseButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.CloseButton.TabIndex = 5;
            this.CloseButton.TabStop = false;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click_1);
            this.CloseButton.MouseLeave += new System.EventHandler(this.CloseButton_MouseLeave_1);
            this.CloseButton.MouseHover += new System.EventHandler(this.CloseButton_MouseHover_1);
            // 
            // Logo
            // 
            this.Logo.BackColor = System.Drawing.Color.Transparent;
            this.Logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.Logo.Image = global::Bobcalc.Properties.Resources.Logo_3;
            this.Logo.Location = new System.Drawing.Point(15, 15);
            this.Logo.Name = "Logo";
            this.Logo.Size = new System.Drawing.Size(140, 142);
            this.Logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.Logo.TabIndex = 4;
            this.Logo.TabStop = false;
            // 
            // Top_Label
            // 
            this.Top_Label.BackColor = System.Drawing.Color.AliceBlue;
            this.Top_Label.Location = new System.Drawing.Point(200, 15);
            this.Top_Label.Name = "Top_Label";
            this.Top_Label.Size = new System.Drawing.Size(820, 63);
            this.Top_Label.TabIndex = 7;
            this.Top_Label.MouseMove += new System.Windows.Forms.MouseEventHandler(this.Top_Label_MouseMove_1);
            // 
            // Start_Button
            // 
            this.Start_Button.BackColor = System.Drawing.Color.AliceBlue;
            this.Start_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Start_Button.Location = new System.Drawing.Point(410, 435);
            this.Start_Button.Name = "Start_Button";
            this.Start_Button.Size = new System.Drawing.Size(220, 60);
            this.Start_Button.TabIndex = 8;
            this.Start_Button.Text = "Начать работу";
            this.Start_Button.UseVisualStyleBackColor = false;
            this.Start_Button.Click += new System.EventHandler(this.Start_Button_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem_Settings});
            this.menuStrip1.Location = new System.Drawing.Point(318, 35);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(199, 29);
            this.menuStrip1.TabIndex = 25;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem_Settings
            // 
            this.toolStripMenuItem_Settings.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.путьДляExcelToolStripMenuItem,
            this.путьДляPDFToolStripMenuItem});
            this.toolStripMenuItem_Settings.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripMenuItem_Settings.Name = "toolStripMenuItem_Settings";
            this.toolStripMenuItem_Settings.Size = new System.Drawing.Size(99, 25);
            this.toolStripMenuItem_Settings.Text = "Настройки";
            // 
            // путьДляExcelToolStripMenuItem
            // 
            this.путьДляExcelToolStripMenuItem.Name = "путьДляExcelToolStripMenuItem";
            this.путьДляExcelToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.путьДляExcelToolStripMenuItem.Text = "Путь для Excel";
            this.путьДляExcelToolStripMenuItem.Click += new System.EventHandler(this.путьДляExcelToolStripMenuItem_Click);
            // 
            // путьДляPDFToolStripMenuItem
            // 
            this.путьДляPDFToolStripMenuItem.Name = "путьДляPDFToolStripMenuItem";
            this.путьДляPDFToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.путьДляPDFToolStripMenuItem.Text = "Путь для PDF";
            this.путьДляPDFToolStripMenuItem.Click += new System.EventHandler(this.путьДляPDFToolStripMenuItem_Click);
            // 
            // Page1
            // 
            this.AcceptButton = this.Start_Button;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(1024, 700);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.Start_Button);
            this.Controls.Add(this.MuteButton);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.Logo);
            this.Controls.Add(this.Top_Label);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Page1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Bob Calc";
            ((System.ComponentModel.ISupportInitialize)(this.MuteButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CloseButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Logo)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox MuteButton;
        private System.Windows.Forms.PictureBox CloseButton;
        private System.Windows.Forms.PictureBox Logo;
        private System.Windows.Forms.Label Top_Label;
        private System.Windows.Forms.Button Start_Button;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem_Settings;
        private System.Windows.Forms.ToolStripMenuItem путьДляExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem путьДляPDFToolStripMenuItem;
    }
}