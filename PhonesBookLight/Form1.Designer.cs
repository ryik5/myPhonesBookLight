namespace PhonesBookLight
{
    partial class FormPhonesBook1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormPhonesBook1));
            this.textBoxData = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.buttonShowAll = new System.Windows.Forms.Button();
            this.Menu1 = new System.Windows.Forms.MenuStrip();
            this.GetDataMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LoadFromServerItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SelectListMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListPhonesItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListFioItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListNavItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListOrgItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SettingsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SetUpItem = new System.Windows.Forms.ToolStripMenuItem();
            this.FolderItem = new System.Windows.Forms.ToolStripMenuItem();
            this.HelpItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.SplitButton1 = new System.Windows.Forms.ToolStripSplitButton();
            this.StatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.SplitButton2 = new System.Windows.Forms.ToolStripSplitButton();
            this.StatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.comboBoxData = new System.Windows.Forms.ComboBox();
            this.panelView = new System.Windows.Forms.Panel();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonSave = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.Menu1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panelView.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxData
            // 
            this.textBoxData.Location = new System.Drawing.Point(390, 28);
            this.textBoxData.Name = "textBoxData";
            this.textBoxData.Size = new System.Drawing.Size(127, 20);
            this.textBoxData.TabIndex = 0;
            this.textBoxData.Click += new System.EventHandler(this.textBoxData_Click);
            this.textBoxData.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBoxData_KeyUp);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.ErrorImage = null;
            this.pictureBox1.Image = global::PhonesBookLight.Properties.Resources.LogoRYIK;
            this.pictureBox1.InitialImage = global::PhonesBookLight.Properties.Resources.LogoRYIK;
            this.pictureBox1.Location = new System.Drawing.Point(522, 57);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(222, 222);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // buttonShowAll
            // 
            this.buttonShowAll.Location = new System.Drawing.Point(8, 27);
            this.buttonShowAll.Name = "buttonShowAll";
            this.buttonShowAll.Size = new System.Drawing.Size(115, 23);
            this.buttonShowAll.TabIndex = 2;
            this.buttonShowAll.Text = "Отобразить все";
            this.buttonShowAll.UseVisualStyleBackColor = true;
            this.buttonShowAll.Click += new System.EventHandler(this.buttonShowAll_Click);
            // 
            // Menu1
            // 
            this.Menu1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.GetDataMenuItem,
            this.SettingsMenuItem,
            this.HelpItem});
            this.Menu1.Location = new System.Drawing.Point(0, 0);
            this.Menu1.Name = "Menu1";
            this.Menu1.Size = new System.Drawing.Size(754, 24);
            this.Menu1.TabIndex = 3;
            this.Menu1.Text = "menuStrip1";
            // 
            // GetDataMenuItem
            // 
            this.GetDataMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.LoadFromServerItem,
            this.SelectListMenuItem,
            this.exitToolStripMenuItem});
            this.GetDataMenuItem.Name = "GetDataMenuItem";
            this.GetDataMenuItem.Size = new System.Drawing.Size(74, 20);
            this.GetDataMenuItem.Text = "Основное";
            // 
            // LoadFromServerItem
            // 
            this.LoadFromServerItem.Name = "LoadFromServerItem";
            this.LoadFromServerItem.Size = new System.Drawing.Size(226, 22);
            this.LoadFromServerItem.Text = "Загрузить список с сервера";
            this.LoadFromServerItem.Click += new System.EventHandler(this.LoadFromServerItem_Click);
            // 
            // SelectListMenuItem
            // 
            this.SelectListMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ListPhonesItem,
            this.ListFioItem,
            this.ListNavItem,
            this.ListOrgItem});
            this.SelectListMenuItem.Name = "SelectListMenuItem";
            this.SelectListMenuItem.Size = new System.Drawing.Size(226, 22);
            this.SelectListMenuItem.Text = "Построить список";
            // 
            // ListPhonesItem
            // 
            this.ListPhonesItem.Name = "ListPhonesItem";
            this.ListPhonesItem.Size = new System.Drawing.Size(159, 22);
            this.ListPhonesItem.Text = "Телефоны";
            this.ListPhonesItem.Click += new System.EventHandler(this.ListPhonesItem_Click);
            // 
            // ListFioItem
            // 
            this.ListFioItem.Name = "ListFioItem";
            this.ListFioItem.Size = new System.Drawing.Size(159, 22);
            this.ListFioItem.Text = "ФИО";
            this.ListFioItem.Click += new System.EventHandler(this.ListFioItem_Click);
            // 
            // ListNavItem
            // 
            this.ListNavItem.Name = "ListNavItem";
            this.ListNavItem.Size = new System.Drawing.Size(159, 22);
            this.ListNavItem.Text = "NAV-коды";
            this.ListNavItem.Click += new System.EventHandler(this.ListNavItem_Click);
            // 
            // ListOrgItem
            // 
            this.ListOrgItem.Name = "ListOrgItem";
            this.ListOrgItem.Size = new System.Drawing.Size(159, 22);
            this.ListOrgItem.Text = "Подразделения";
            this.ListOrgItem.Click += new System.EventHandler(this.ListOrgItem_Click);
            // 
            // SettingsMenuItem
            // 
            this.SettingsMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.SetUpItem,
            this.FolderItem});
            this.SettingsMenuItem.Name = "SettingsMenuItem";
            this.SettingsMenuItem.Size = new System.Drawing.Size(79, 20);
            this.SettingsMenuItem.Text = "Настройки";
            // 
            // SetUpItem
            // 
            this.SetUpItem.Name = "SetUpItem";
            this.SetUpItem.Size = new System.Drawing.Size(203, 22);
            this.SetUpItem.Text = "Настройки программы";
            this.SetUpItem.Click += new System.EventHandler(this.SetUpItem_Click);
            // 
            // FolderItem
            // 
            this.FolderItem.Name = "FolderItem";
            this.FolderItem.Size = new System.Drawing.Size(203, 22);
            this.FolderItem.Text = "Папка программы";
            this.FolderItem.Click += new System.EventHandler(this.FolderItem_Click);
            // 
            // HelpItem
            // 
            this.HelpItem.Name = "HelpItem";
            this.HelpItem.Size = new System.Drawing.Size(68, 20);
            this.HelpItem.Text = "Помощь";
            this.HelpItem.Click += new System.EventHandler(this.HelpItem_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusLabel1,
            this.SplitButton1,
            this.StatusLabel2,
            this.SplitButton2,
            this.StatusLabel3,
            this.progressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 486);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(754, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // StatusLabel1
            // 
            this.StatusLabel1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.StatusLabel1.Name = "StatusLabel1";
            this.StatusLabel1.Size = new System.Drawing.Size(73, 17);
            this.StatusLabel1.Text = "StatusLabel1";
            // 
            // SplitButton1
            // 
            this.SplitButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.SplitButton1.Image = global::PhonesBookLight.Properties.Resources.LogoRYIK;
            this.SplitButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.SplitButton1.Name = "SplitButton1";
            this.SplitButton1.Size = new System.Drawing.Size(32, 20);
            this.SplitButton1.Text = "SplitButton1";
            // 
            // StatusLabel2
            // 
            this.StatusLabel2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.StatusLabel2.Name = "StatusLabel2";
            this.StatusLabel2.Size = new System.Drawing.Size(73, 17);
            this.StatusLabel2.Text = "StatusLabel2";
            this.StatusLabel2.TextChanged += new System.EventHandler(this.StatusLabel2_TextChanged);
            // 
            // SplitButton2
            // 
            this.SplitButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.SplitButton2.Image = global::PhonesBookLight.Properties.Resources.LogoRYIK;
            this.SplitButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.SplitButton2.Name = "SplitButton2";
            this.SplitButton2.Size = new System.Drawing.Size(32, 20);
            this.SplitButton2.Text = "SplitButton2";
            // 
            // StatusLabel3
            // 
            this.StatusLabel3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.StatusLabel3.Name = "StatusLabel3";
            this.StatusLabel3.Size = new System.Drawing.Size(427, 17);
            this.StatusLabel3.Spring = true;
            this.StatusLabel3.Text = "StatusLabel3";
            this.StatusLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.StatusLabel3.TextChanged += new System.EventHandler(this.StatusLabel3_TextChanged);
            // 
            // progressBar1
            // 
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(8, 57);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(509, 426);
            this.dataGridView1.TabIndex = 5;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_DoubleClick);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Visible = true;
            // 
            // comboBoxData
            // 
            this.comboBoxData.FormattingEnabled = true;
            this.comboBoxData.Location = new System.Drawing.Point(132, 28);
            this.comboBoxData.Name = "comboBoxData";
            this.comboBoxData.Size = new System.Drawing.Size(252, 21);
            this.comboBoxData.TabIndex = 6;
            this.comboBoxData.SelectedIndexChanged += new System.EventHandler(this.comboBoxData_SelectedIndexChanged);
            // 
            // panelView
            // 
            this.panelView.Controls.Add(this.buttonCancel);
            this.panelView.Controls.Add(this.buttonSave);
            this.panelView.Location = new System.Drawing.Point(3, 28);
            this.panelView.Name = "panelView";
            this.panelView.Size = new System.Drawing.Size(746, 455);
            this.panelView.TabIndex = 7;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(190, 14);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(100, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Отменить";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(20, 14);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(100, 23);
            this.buttonSave.TabIndex = 0;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(226, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // FormPhonesBook1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Azure;
            this.ClientSize = new System.Drawing.Size(754, 508);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.comboBoxData);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.buttonShowAll);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.textBoxData);
            this.Controls.Add(this.Menu1);
            this.Controls.Add(this.panelView);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormPhonesBook1";
            this.Text = "PhonesBook";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormPhonesBook1_FormClosed);
            this.Load += new System.EventHandler(this.FormPhonesBook1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.Menu1.ResumeLayout(false);
            this.Menu1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panelView.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxData;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button buttonShowAll;
        private System.Windows.Forms.MenuStrip Menu1;
        private System.Windows.Forms.ToolStripMenuItem GetDataMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LoadFromServerItem;
        private System.Windows.Forms.ToolStripMenuItem SelectListMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ComboBox comboBoxData;
        private System.Windows.Forms.ToolStripMenuItem SettingsMenuItem;
        private System.Windows.Forms.Panel panelView;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.ToolStripSplitButton SplitButton1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem ListPhonesItem;
        private System.Windows.Forms.ToolStripMenuItem ListFioItem;
        private System.Windows.Forms.ToolStripMenuItem ListNavItem;
        private System.Windows.Forms.ToolStripSplitButton SplitButton2;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel3;
        private System.Windows.Forms.ToolStripMenuItem ListOrgItem;
        private System.Windows.Forms.ToolStripMenuItem FolderItem;
        private System.Windows.Forms.ToolStripMenuItem HelpItem;
        private System.Windows.Forms.ToolStripMenuItem SetUpItem;
        private System.Windows.Forms.ToolStripProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
    }
}

