namespace Geochemistry
{
    partial class frmPpal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPpal));
            this.mnuPpal = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.passwordChangeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logOutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.geochemistryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.soilToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rockToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sedimentToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.channelsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuMinnGeol = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.mnuPpal.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnuPpal
            // 
            this.mnuPpal.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.geochemistryToolStripMenuItem});
            this.mnuPpal.Location = new System.Drawing.Point(0, 0);
            this.mnuPpal.Name = "mnuPpal";
            this.mnuPpal.Size = new System.Drawing.Size(776, 24);
            this.mnuPpal.TabIndex = 1;
            this.mnuPpal.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.passwordChangeToolStripMenuItem,
            this.logOutToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // passwordChangeToolStripMenuItem
            // 
            this.passwordChangeToolStripMenuItem.Name = "passwordChangeToolStripMenuItem";
            this.passwordChangeToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.passwordChangeToolStripMenuItem.Text = "Password Change";
            this.passwordChangeToolStripMenuItem.Click += new System.EventHandler(this.passwordChangeToolStripMenuItem_Click);
            // 
            // logOutToolStripMenuItem
            // 
            this.logOutToolStripMenuItem.Name = "logOutToolStripMenuItem";
            this.logOutToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.logOutToolStripMenuItem.Text = "Log out";
            this.logOutToolStripMenuItem.Click += new System.EventHandler(this.logOutToolStripMenuItem_Click);
            // 
            // geochemistryToolStripMenuItem
            // 
            this.geochemistryToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.soilToolStripMenuItem,
            this.rockToolStripMenuItem,
            this.sedimentToolStripMenuItem,
            this.channelsToolStripMenuItem,
            this.MenuMinnGeol});
            this.geochemistryToolStripMenuItem.Name = "geochemistryToolStripMenuItem";
            this.geochemistryToolStripMenuItem.Size = new System.Drawing.Size(92, 20);
            this.geochemistryToolStripMenuItem.Text = "Geochemistry";
            // 
            // soilToolStripMenuItem
            // 
            this.soilToolStripMenuItem.Name = "soilToolStripMenuItem";
            this.soilToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.soilToolStripMenuItem.Text = "Soil";
            this.soilToolStripMenuItem.Visible = false;
            this.soilToolStripMenuItem.Click += new System.EventHandler(this.soilToolStripMenuItem_Click);
            // 
            // rockToolStripMenuItem
            // 
            this.rockToolStripMenuItem.Name = "rockToolStripMenuItem";
            this.rockToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.rockToolStripMenuItem.Text = "Rock";
            this.rockToolStripMenuItem.Click += new System.EventHandler(this.rockToolStripMenuItem_Click);
            // 
            // sedimentToolStripMenuItem
            // 
            this.sedimentToolStripMenuItem.Name = "sedimentToolStripMenuItem";
            this.sedimentToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.sedimentToolStripMenuItem.Text = "Sediment";
            this.sedimentToolStripMenuItem.Visible = false;
            this.sedimentToolStripMenuItem.Click += new System.EventHandler(this.sedimentToolStripMenuItem_Click);
            // 
            // channelsToolStripMenuItem
            // 
            this.channelsToolStripMenuItem.Name = "channelsToolStripMenuItem";
            this.channelsToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.channelsToolStripMenuItem.Text = "Channels";
            this.channelsToolStripMenuItem.Click += new System.EventHandler(this.channelsToolStripMenuItem_Click);
            // 
            // MenuMinnGeol
            // 
            this.MenuMinnGeol.Name = "MenuMinnGeol";
            this.MenuMinnGeol.Size = new System.Drawing.Size(180, 22);
            this.MenuMinnGeol.Text = "Minning Geology";
            this.MenuMinnGeol.Click += new System.EventHandler(this.MenuMinnGeol_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 489);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(776, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // frmPpal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(776, 511);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.mnuPpal);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.mnuPpal;
            this.Name = "frmPpal";
            this.Text = "GeoChemistry";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmPpal_FormClosed);
            this.Load += new System.EventHandler(this.frmPpal_Load);
            this.mnuPpal.ResumeLayout(false);
            this.mnuPpal.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip mnuPpal;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem passwordChangeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem logOutToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripMenuItem geochemistryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem soilToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rockToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sedimentToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem channelsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem MenuMinnGeol;
    }
}