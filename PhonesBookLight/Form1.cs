using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace PhonesBookLight
{
    public partial class FormPhonesBook1 : Form
    {
        private System.Diagnostics.FileVersionInfo myFileVersionInfo;
        private System.Windows.Forms.ContextMenu contextMenu1;

        private string UserWindowsAuthorization = ";Persist Security Info=True";
        private string ServerName;
        private string UserLogin;
        private string UserPassword;
        private string sFolderPhotos = @".\Photos\";
        private readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        private readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        private string myRegKey = @"SOFTWARE\RYIK\PhonesBookLight";
        private int iRowRecords = 0;
        private int iRowFIO = 0;
        private HashSet<string> lData = new HashSet<string>();
        private string sLastGotData = "";

        private DataTable dtPeriod = new DataTable("PeriodListData");
        private DataColumn[] dcPeriod ={
                                  new DataColumn("№ п/п",typeof(double)),
                                  new DataColumn("Номер телефона",typeof(string)),
                                  new DataColumn("ФИО",typeof(string)),
                                  new DataColumn("NAV",typeof(string)),
                                  new DataColumn("Подразделение",typeof(string)),
                                  new DataColumn("Основной",typeof(string)),
                                  new DataColumn("Действует c",typeof(string)),
                                  new DataColumn("Действует по",typeof(string)),
                              };
        private DataColumn[] dcPeriodkeys = new DataColumn[2];
        private DataTable dtPeriodShow = new DataTable("PeriodListData");
        private DataColumn[] dcPeriodShow ={
                                  new DataColumn("№ п/п",typeof(double)),
                                  new DataColumn("Номер телефона",typeof(string)),
                                  new DataColumn("ФИО",typeof(string)),
                                  new DataColumn("NAV",typeof(string)),
                                  new DataColumn("Подразделение",typeof(string)),
                                  new DataColumn("Основной",typeof(string)),
                                  new DataColumn("Действует c",typeof(string)),
                                  new DataColumn("Действует по",typeof(string)),
                              };
        private DataColumn[] dcPeriodShowkeys = new DataColumn[1];

        private Label labelServer;
        private TextBox textBoxServer;
        private Label labelServerUserName;
        private TextBox textBoxServerUserName;
        private Label labelServerUserPassword;
        private TextBox textBoxServerUserPassword;
        private Label labelFolderPhotos;
        private TextBox textFolderPhotos;
        private ToolTip toolTip1 = new ToolTip();
        private Bitmap bmp;
        private string sSelectedNAV;
        private string sSelectedFIO;
        private string sSelectedPhone;
        private string sSelected4;
        private string sSelected5;
        private string sSelected6;
        private Label lString1;
        private Label lString2;
        private Label lString3;
        private Label lString4;
        private Label lString5;
        private Label lString6;

        public FormPhonesBook1()
        { InitializeComponent(); }

        private void FormPhonesBook1_Load(object sender, EventArgs e)
        { LoadForm(); }

        private void LoadForm()
        {
            buttonShowAll.Enabled = false;
            panelView.Visible = false;
            dataGridView1.Visible = true;
            dataGridView1.BringToFront();

            bmp = new Bitmap(Properties.Resources.LogoRYIK, pictureBox1.Width, pictureBox1.Height);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.BorderStyle = BorderStyle.None;
            RefreshPictureBox(pictureBox1, bmp);

            panelViewResize();
            CheckRegistrySavedData();

            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            StatusLabel2.Text = "";
            StatusLabel3.Text = "";
            toolTip1.SetToolTip(textBoxData, "Поле ввода данных для поиска.\nЭто может быть ФИО полностью или его часть\n" +
            "Это может быть весь номер или же часть его.\nЭто может быть NAV-код или часть наименования организации");
            LoadFromServerItem.ToolTipText = "Загрузить список ФИО/телефонов с сервера T-factura";
            SetUpItem.ToolTipText = "Внести настройки в программу,\nнеобходимые для выполнения корректного подключения и авторизации на сервере Т-Factura.\n" +
                "Все настройки хранятся в профиле пользователя в зашифрованном виде.";
            SelectListMenuItem.Enabled = false;


            dtPeriod.Columns.AddRange(dcPeriod);
            dcPeriodkeys[0] = dcPeriod[1];
            dcPeriodkeys[1] = dcPeriod[2];
            dtPeriod.PrimaryKey = dcPeriodkeys;
            dtPeriodShow.Columns.AddRange(dcPeriodShow);
            dcPeriodShowkeys[0] = dcPeriodShow[0];
            dtPeriodShow.PrimaryKey = dcPeriodShowkeys;
            contextMenu1 = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon1.ContextMenu = contextMenu1;
            contextMenu1.MenuItems.Add("About", AboutSoft);
            contextMenu1.MenuItems.Add("Exit", ApplicationExit);
            notifyIcon1.Text = myFileVersionInfo.ProductName + "\nv." + myFileVersionInfo.FileVersion + "\n" + myFileVersionInfo.CompanyName;
            this.Text = myFileVersionInfo.Comments;
        }

        private void AboutSoft(object sender, EventArgs e) //for the Context Menu on notify Icon
        { AboutSoft(); }

        private void ApplicationExit(object sender, EventArgs e) //for the Context Menu on notify Icon
        { ApplicationExit(); }

        private void ApplicationExit()
        { Application.Exit(); }

        private void AboutSoft()
        {
            String strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            DialogResult result = MessageBox.Show(
               myFileVersionInfo.Comments + "\n" + "Версия: " + myFileVersionInfo.FileVersion + "\nBuild: " +
                strVersion + "\n" + myFileVersionInfo.LegalCopyright +
                "\n\n1. Первый раз, перед получением данных с сервера T-Factura, необходимо:\n" + @"    A) Ввести адрес сервера в виде - SERVER.DOMAIN.SUBDOMAIN" + "\n" +
                @"    B) Ввести авторизационные данные" +
                "\n" + @"    С) Нажать кнопку " + "\"Сохранить\".\n" +
                "2. Корректный адрес сервера, имя и пароль пользователя T-Factura, можно получить в ИТ - отделе.\n" +                
                "\nOriginal file: " + myFileVersionInfo.OriginalFilename + "\nFull path: " + Application.ExecutablePath,
                "Информация об использовании программы",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }


        private void CheckRegistrySavedData()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(myRegKey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey))
                {
                    ServerName = DecryptBase64ToString(EvUserKey.GetValue("ServerName").ToString(), btsMess1, btsMess2);
                    UserLogin = DecryptBase64ToString(EvUserKey.GetValue("UserLogin").ToString(), btsMess1, btsMess2);
                    UserPassword = DecryptBase64ToString(EvUserKey.GetValue("UserPassword").ToString(), btsMess1, btsMess2);
                    sFolderPhotos = EvUserKey.GetValue("FolderPhotos").ToString();
                }
            }
            catch { }
        }

        private void ListPhonesItem_Click(object sender, EventArgs e)
        { MakeListData("phone_no"); }

        private void ListFioItem_Click(object sender, EventArgs e)
        { MakeListData("emp_name"); }

        private void ListNavItem_Click(object sender, EventArgs e)
        { MakeListData("NAV"); }

        private void ListOrgItem_Click(object sender, EventArgs e)
        { MakeListData("org_unit_name"); }

        private void LoadFromServerItem_Click(object sender, EventArgs e) //use
        { LoadFromServer(); }

        private void LoadFromServer()
        {
            buttonShowAll.Enabled = false;
            sLastGotData = "LoadFromServer";
            GetDataFromServer();
            if (dtPeriod.Rows.Count > 1)
            {
                StatusLabel3.Text = "Обрабатываю полученные данные...";
                StatusLabel3.ForeColor = Color.Black;
                SelectListMenuItem.Enabled = true;
                MakeListData("emp_name");
                MakeTable();
                StatusLabel2.Text = "Всего ФИО - " + iRowFIO.ToString() + " | Всего номеров - " + dtPeriod.Rows.Count.ToString();
                StatusLabel3.Text = "Готово!";
                dataGridView1.Columns[7].Visible = false;
            }
            else
            {
                StatusLabel3.Text = "Данные с выбранного сервера не получены!";
                StatusLabel3.ForeColor = Color.DarkRed;
            }
        }

        private void GetDataFromServer()
        {
            try
            {
                string sSqlQuery;
                string sConnection = @"Data Source=" + ServerName + @";Initial Catalog=EBP;Type System Version=SQL Server 2005" + UserWindowsAuthorization + @";User ID=" + UserLogin + @";Password=" + UserPassword + @";Connect Timeout=60";
                using (System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(sConnection))
                {
                    sqlConnection.Open();
                    sSqlQuery =
                        "SELECT t1.phone_no AS phone_no, t1.emp_name AS emp_name, t1.org_unit_name AS org_unit_name, t1.till_dt AS till_dt, t1.from_dt as from_dt, t1.descr AS main, os_emp.emp_cd AS NAV " +
                        "FROM v_rs_contract_detail t1 INNER JOIN os_emp ON os_emp.emp_id=t1.emp_id " +
                        "WHERE t1.till_dt is null AND emp_name is not null AND state LIKE '%A%' AND phone_no NOT LIKE '395%' AND emp_name NOT LIKE '%резерв%' AND emp_name NOT LIKE '%шлюз%' AND os_emp.emp_cd NOT LIKE 'S00385' AND os_emp.emp_cd NOT LIKE 'S01557' " +
                        "ORDER by emp_name, phone_no";

                    dtPeriod.Rows.Clear();
                    iRowRecords = 0;

                    using (System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(sSqlQuery, sqlConnection))
                    {
                        using (System.Data.SqlClient.SqlDataReader sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (System.Data.Common.DbDataRecord record in sqlReader)
                            {
                                if (record != null && record.ToString().Length > 0 && record["phone_no"].ToString().Length > 0)
                                {
                                    DataRow row = dtPeriod.NewRow();
                                    row["№ п/п"] = ++iRowRecords;
                                    row["Номер телефона"] = MakeCommonViewPhone(record["phone_no"].ToString());
                                    row["ФИО"] = record["emp_name"].ToString().Trim();
                                    row["NAV"] = record["NAV"].ToString().Trim();
                                    row["Подразделение"] = record["org_unit_name"].ToString().Trim();
                                    row["Основной"] = DefineMainPhone(record["main"].ToString());
                                    row["Действует c"] = record["from_dt"].ToString().Trim().Split(' ')[0];
                                    dtPeriod.Rows.Add(row);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }
        }

        private void MakeListData(string sFindData)
        {
            lData = new HashSet<string>();
            List<string> l = new List<string>();
            string sCellFound = "";

            switch (sFindData.ToLower())
            {
                case ("phone_no"):
                    sCellFound = "Номер телефона";
                    StatusLabel3.Text = "Сгенерирован список телефонов";
                    break;
                case ("номер телефона"):
                    sCellFound = "Номер телефона";
                    StatusLabel3.Text = "Сгенерирован список телефонов";
                    break;
                case ("emp_name"):
                    sCellFound = "ФИО";
                    StatusLabel3.Text = "Сгенерирован список ФИО";
                    break;
                case ("фио"):
                    sCellFound = "ФИО";
                    StatusLabel3.Text = "Сгенерирован список ФИО";
                    break;
                case ("nav"):
                    sCellFound = "NAV";
                    StatusLabel3.Text = "Сгенерирован список NAV-кодов";
                    break;
                case ("org_unit_name"):
                    sCellFound = "Подразделение";
                    StatusLabel3.Text = "Сгенерирован список подразделений";
                    break;
                case ("подразделение"):
                    sCellFound = "Подразделение";
                    StatusLabel3.Text = "Сгенерирован список подразделений";
                    break;
            }

            if (sLastGotData == "LoadFromServer")
            { TableSearchToLdata(dtPeriod, lData, sCellFound); }

            if (lData.Count > 0)
            {
                iRowFIO = lData.Count;
                try
                {
                    if (this.InvokeRequired)
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            comboBoxData.Items.Clear();
                            comboBoxData.Sorted = true;
                            AutoCompleteStringCollection sourceList = new AutoCompleteStringCollection();
                            sourceList.AddRange(lData.ToArray());
                            comboBoxData.Items.AddRange(lData.ToArray());
                            comboBoxData.AutoCompleteCustomSource = sourceList;
                            comboBoxData.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                            comboBoxData.AutoCompleteSource = AutoCompleteSource.CustomSource;
                        }));
                    else
                    {
                        comboBoxData.Items.Clear();
                        comboBoxData.Sorted = true;
                        AutoCompleteStringCollection sourceList = new AutoCompleteStringCollection();
                        sourceList.AddRange(lData.ToArray());
                        comboBoxData.Items.AddRange(lData.ToArray());
                        comboBoxData.AutoCompleteCustomSource = sourceList;
                        comboBoxData.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                        comboBoxData.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    }
                }
                catch { }
            }
            else { iRowFIO = 0; }

            StatusLabel2.Text = "Всего записей: " + iRowFIO.ToString();
        }

        private void SetUpItem_Click(object sender, EventArgs e)
        { ShowSettingsData(); }

        private void ShowSettingsData() //Perform of Controls Settings Data
        {
            dataGridView1.Visible = false;
            pictureBox1.Visible = false;
            comboBoxData.Visible = false;
            textBoxData.Visible = false;
            buttonShowAll.Visible = false;
            GetDataMenuItem.Visible = false;

            panelView.BorderStyle = BorderStyle.FixedSingle;
            panelView.Visible = true;
            panelViewResize();

            labelServer = new Label
            {
                Text = "Server",
                BackColor = Color.PaleGreen,
                Location = new Point(20, 60),
                Size = new Size(590, 22),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = panelView
            };
            textBoxServer = new TextBox
            {
                Text = "",
                PasswordChar = '*',
                Location = new Point(90, 61),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = panelView
            };
            toolTip1.SetToolTip(textBoxServer, "Имя сервера с базой T-factura в виде:\nNameOfServer.Domain.Subdomain");
            textBoxServer.Click += new System.EventHandler(ResettextBoxServer);

            labelServerUserName = new Label
            {
                Text = "UserName",
                BackColor = Color.PaleGreen,
                Location = new Point(220, 61),
                Size = new Size(70, 20),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = panelView
            };
            textBoxServerUserName = new TextBox
            {
                Text = "",
                PasswordChar = '*',
                Location = new Point(300, 61),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = panelView
            };
            toolTip1.SetToolTip(textBoxServerUserName, "Имя пользователя SQL-сервера T-factura");
            textBoxServerUserName.Click += new System.EventHandler(ResetServerUserName);

            labelServerUserPassword = new Label
            {
                Text = "Password",
                BackColor = Color.PaleGreen,
                Location = new Point(420, 61),
                Size = new Size(70, 20),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = panelView
            };
            textBoxServerUserPassword = new TextBox
            {
                Text = "",
                PasswordChar = '*',
                Location = new Point(500, 61),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = panelView
            };
            toolTip1.SetToolTip(textBoxServerUserPassword, "Пароль администратора SQL-сервера T-factura");
            textBoxServerUserPassword.Click += new System.EventHandler(ResetServerUserPassword);

            labelFolderPhotos = new Label
            {
                Text = "Photos",
                BackColor = Color.PaleGreen,
                Location = new Point(20, 90),
                Size = new Size(70, 20),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = panelView
            };
            textFolderPhotos = new TextBox
            {
                Text = sFolderPhotos,
                Location = new Point(90, 91),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = panelView
            };
            textFolderPhotos.Click += new System.EventHandler(ResetFolderPhotos);
            toolTip1.SetToolTip(textFolderPhotos, "Путь к папке с фотографиями.\nМожно указать относительным к папке с программой в виде:\n.\\Photos\\");

            buttonSave.FlatStyle = FlatStyle.Flat;
            buttonCancel.FlatStyle = FlatStyle.Flat;
            buttonSave.FlatAppearance.MouseOverBackColor = Color.PaleGreen; //Change color of button if mouse over the button
            buttonCancel.FlatAppearance.MouseOverBackColor = Color.PaleGreen;

            labelFolderPhotos.BringToFront();
            labelServerUserName.BringToFront();
            labelServerUserPassword.BringToFront();
            textBoxServer.BringToFront();
            textBoxServerUserName.BringToFront();
            textBoxServerUserPassword.BringToFront();
            textFolderPhotos.BringToFront();

            if (UserLogin != null && UserPassword != null && ServerName != null && UserLogin.Length > 0 && UserPassword.Length > 0 && ServerName.Length > 0)
            {
                textBoxServer.Text = ServerName;
                textBoxServerUserName.Text = UserLogin;
                textBoxServerUserPassword.Text = UserPassword;
            }
        }

        private void ResetFolderPhotos(object sender, EventArgs e)
        { textFolderPhotos.Clear(); }

        private void ResetServerUserPassword(object sender, EventArgs e)
        { textBoxServerUserPassword.Clear(); }

        private void ResetServerUserName(object sender, EventArgs e)
        { textBoxServerUserName.Clear(); }

        private void ResettextBoxServer(object sender, EventArgs e)
        { textBoxServer.Clear(); }

        private void buttonCancel_Click(object sender, EventArgs e) //Use Cancel()
        { Cancel(); }

        private void Cancel() //Cancel inputed Data
        {
            panelView.Visible = false;

            labelServer.Dispose();
            textBoxServer.Dispose();
            labelServerUserName.Dispose();
            textBoxServerUserName.Dispose();
            labelServerUserPassword.Dispose();
            textBoxServerUserPassword.Dispose();

            dataGridView1.Visible = true;
            pictureBox1.Visible = true;
            comboBoxData.Visible = true;
            textBoxData.Visible = true;
            buttonShowAll.Visible = true;
            GetDataMenuItem.Visible = true;
        }

        private void buttonSave_Click(object sender, EventArgs e) // Use Save()
        { Save(); }

        private void Save() //Save inputed Credintials and Parameters into Registry and variables
        {
            ServerName = textBoxServer.Text;
            UserLogin = textBoxServerUserName.Text;
            UserPassword = textBoxServerUserPassword.Text;
            sFolderPhotos = textFolderPhotos.Text;

            if (UserLogin != null && UserPassword != null && ServerName != null && UserLogin.Length > 0 && UserPassword.Length > 0 && ServerName.Length > 0)
            {
                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue("ServerName", EncryptStringToBase64Text(ServerName, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("UserLogin", EncryptStringToBase64Text(UserLogin, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("UserPassword", EncryptStringToBase64Text(UserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("FolderPhotos", sFolderPhotos, Microsoft.Win32.RegistryValueKind.String);
                    }
                }
                catch { }
            }

            panelView.Visible = false;

            labelServer.Dispose();
            textBoxServer.Dispose();
            labelServerUserName.Dispose();
            textBoxServerUserName.Dispose();
            labelServerUserPassword.Dispose();
            textBoxServerUserPassword.Dispose();
            labelFolderPhotos.Dispose();
            textFolderPhotos.Dispose();

            dataGridView1.Visible = true;
            pictureBox1.Visible = true;
            comboBoxData.Visible = true;
            textBoxData.Visible = true;
            buttonShowAll.Visible = true;
            GetDataMenuItem.Visible = true;
        }

        private void panelView_SizeChanged(object sender, EventArgs e)
        { panelViewResize(); }

        private void panelViewResize() //Change PanelView
        {
            panelView.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
            panelView.Height = panelView.Parent.Height - 120;
            panelView.AutoScroll = true;
            panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            panelView.ResumeLayout();
        }

        private void RefreshPictureBox(PictureBox picBox, Bitmap picImage) // не работает
        {
            picBox.Image = RefreshBitmap(picImage, panelView.Width - 2, panelView.Height - 2); //сжатая картина
            picBox.Refresh();
        }

        private Bitmap RefreshBitmap(Bitmap b, int nWidth, int nHeight)
        {
            Bitmap result = new Bitmap(nWidth, nHeight);
            using (Graphics g = Graphics.FromImage((Image)result))
            { g.DrawImage(b, 0, 0, nWidth, nHeight); }
            return result;
        }

        private void buttonShowAll_Click(object sender, EventArgs e) //Use ShowAll()
        { ShowAll(); }

        private void ShowAll() //Show all data from lFullData()
        {
            SelectListMenuItem.Enabled = true;
            buttonShowAll.Enabled = false;
            MakeTable();

            textBoxData.Clear();
            comboBoxData.SelectedText = null;

            StatusLabel3.Text = "Фильтр отключен";
            StatusLabel2.Text = "Всего записей - " + iRowRecords.ToString();
        }

        private void MakeTable()
        {
            iRowRecords = 0;
            if (sLastGotData == "LoadFromServer")
            {
                TableToTableshow(dtPeriod, dtPeriodShow);
                dataGridView1.Columns[7].Visible = false;
            }
        }

        private void DatagridCollumnsFullTableToHide(DataGridView dgv, int[] showFullTableCollumns)
        {
            dgv.Columns[8].Visible = showFullTableCollumns[0] > 0 ? true : false;
            dgv.Columns[7].Visible = showFullTableCollumns[10] > 0 ? true : false;
            dgv.Columns[6].Visible = showFullTableCollumns[9] > 0 ? true : false;
            dgv.Columns[2].Visible = showFullTableCollumns[8] > 0 ? true : false;
            dgv.Columns[5].Visible = showFullTableCollumns[5] > 0 ? true : false;
            dgv.Columns[4].Visible = showFullTableCollumns[2] > 0 ? true : false;
        }

        private static string EncryptStringToBase64Text(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data. Use EncryptStringToBytes()
        {
            string sBase64Test;
            sBase64Test = Convert.ToBase64String(EncryptStringToBytes(plainText, Key, IV));
            return sBase64Test;
        }

        private static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;

            using (RijndaelManaged rijAlg = new RijndaelManaged())        // Create an RijndaelManaged object with the specified key and IV. 
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);    // Create a decrytor to perform the stream transform.

                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())   // Create the streams used for encryption. 
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);   //Write all data to the stream.
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            return encrypted;    // Return the encrypted bytes from the memory stream. 
        }

        private static string DecryptBase64ToString(string sBase64Text, byte[] Key, byte[] IV) //Encrypt variables PlainText Data. Use DecryptStringFromBytes()
        {
            byte[] bBase64Test;
            bBase64Test = Convert.FromBase64String(sBase64Text);
            return DecryptStringFromBytes(bBase64Test, Key, IV);
        }

        private static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV) //Decrypt PlainText Data to variables
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            string plaintext = null;   // Declare the string used to hold the decrypted text.

            using (RijndaelManaged rijAlg = new RijndaelManaged())  // Create an RijndaelManaged object  with the specified key and IV.
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);  // Create a decrytor to perform the stream transform.                              

                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))  // Create the streams used for decryption. 
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
                        {
                            plaintext = srDecrypt.ReadToEnd();  // Read the decrypted bytes from the decrypting stream and place them in a string. 
                        }
                    }
                }
            }
            return plaintext;
        }

        private void FormPhonesBook1_FormClosed(object sender, FormClosedEventArgs e)
        { Application.Exit(); }

        private string MakeCommonViewPhone(string sPrimaryPhone) //Normalize Phone to +380504197443
        {
            string sPhone = sPrimaryPhone.Trim();
            string sTemp1 = "", sTemp2 = "";
            sTemp1 = sPhone.Replace(" ", "");
            sTemp2 = sTemp1.Replace("-", "");
            sTemp1 = sTemp2.Replace(")", "");
            sTemp2 = sTemp1.Replace("(", "");
            sTemp1 = sTemp2.Replace("/", "");
            sTemp2 = sTemp1.Replace("_", "");

            if (sTemp2.StartsWith("+") && sTemp2.Length == 13) sPhone = sTemp2;
            else if (sTemp2.StartsWith("380") && sTemp2.Length == 12) sPhone = "+" + sTemp2;
            else if (sTemp2.StartsWith("80") && sTemp2.Length == 11) sPhone = "+3" + sTemp2;
            else if (sTemp2.StartsWith("0") && sTemp2.Length == 10) sPhone = "+38" + sTemp2;
            else if (sTemp2.Length == 9) sPhone = "+380" + sTemp2;
            else sPhone = sTemp2;

            sTemp1 = ""; sTemp2 = "";
            return sPhone;
        }

        private string DefineMainPhone(string sDescription)
        {
            if (sDescription.Trim() == "!")
            { return "Да"; }
            else { return ""; }
        }

        private void StatusLabel2_TextChanged(object sender, EventArgs e)
        {
            if (StatusLabel2.Text.Length > 0)
            { SplitButton1.Visible = true; }
            else
            { SplitButton1.Visible = false; }
        }

        private void StatusLabel3_TextChanged(object sender, EventArgs e)
        {
            if (StatusLabel3.Text.Length > 0)
            { SplitButton2.Visible = true; }
            else
            { SplitButton2.Visible = false; }
        }

        private void textBoxData_KeyUp(object sender, KeyEventArgs e) //Action after pressed the button "Enter"
        {
            if (e.KeyCode == Keys.Enter)
            {
                string sSelected = textBoxData.Text.Trim();
                DataSearch(sSelected);
                StatusLabel3.Text = "Фильтр включен";
            }
            StatusLabel2.Text = "Найдено записей - " + iRowRecords;
        }

        private void comboBoxData_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sSelected = comboBoxData.SelectedItem.ToString();
            DataSearch(sSelected);
        }

        private void DataSearch(string sSelected)
        {
            SearchSelectedData(sSelected);
            ShowData(sSelectedPhone, sSelectedFIO, sSelectedNAV, sSelected4, sSelected5, sSelected6);
            buttonShowAll.Enabled = true;
        }

        private void SearchSelectedData(string sSelected)
        {
            iRowRecords = 0;
            sSelectedPhone = "";
            sSelectedFIO = "";
            sSelectedNAV = "";
            sSelected4 = "";
            sSelected5 = "";
            sSelected6 = "";

            lString1?.Dispose();
            lString2?.Dispose();
            lString3?.Dispose();
            lString4?.Dispose();
            lString5?.Dispose();
            lString6?.Dispose();

            if (sLastGotData == "LoadFromServer")
            { TableSearchToTableshow(dtPeriod, dtPeriodShow, sSelected); }
        }

        private void wc_DownloadProgressChanged(object sender, System.Net.DownloadProgressChangedEventArgs e)
        { progressBar1.Value = e.ProgressPercentage; }

        private void ShowData(string s1 = "", string s2 = "", string s3 = "", string s4 = "", string s5 = "", string s6 = "") //Show the Loaded Picture in The PictureBox1
        {
            string photoPath = "";
            System.IO.FileInfo photoFileInfo = new System.IO.FileInfo("new.jpg");

            try
            {
                progressBar1.Value = 0;
                pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                photoPath = sFolderPhotos + sSelectedNAV + @".jpg";

                photoFileInfo = new System.IO.FileInfo(photoPath);
                if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                { System.IO.File.Delete(photoPath); }

                pictureBox1.Load(photoPath);
            }
            catch
            {
                string sTempNav = sSelectedNAV;
                if (sSelectedNAV.Contains("С")) //russkaya С                       
                {
                    sTempNav = sSelectedNAV.Replace("С", "S");
                    try
                    {
                        pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                        photoPath = sFolderPhotos + sTempNav + @".jpg";

                        photoFileInfo = new System.IO.FileInfo(photoPath);
                        if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                        { System.IO.File.Delete(photoPath); }

                        pictureBox1.Load(photoPath);
                    }
                    catch
                    {
                        try
                        {
                            sTempNav = sSelectedNAV.Replace("С", "C");//english С
                            pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                            photoPath = sFolderPhotos + sTempNav + @".jpg";

                            photoFileInfo = new System.IO.FileInfo(photoPath);
                            if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                            { System.IO.File.Delete(photoPath); }

                            pictureBox1.Load(photoPath);
                        }
                        catch
                        {
                            try
                            {
                                try { System.IO.Directory.CreateDirectory(sFolderPhotos); } catch { }
                                progressBar1.Value = 0;
                                using (System.Net.WebClient wc = new System.Net.WebClient())
                                {
                                    wc.DownloadProgressChanged += wc_DownloadProgressChanged;
                                    string url = "http://www.ais/usersimage/Fotos/" + sSelectedNAV + @".jpg";
                                    string save_path = sFolderPhotos + sSelectedNAV + @".jpg";
                                    wc.DownloadFileAsync(new Uri(url), save_path);
                                    pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                                    photoPath = sFolderPhotos + sSelectedNAV + @".jpg";

                                    photoFileInfo = new System.IO.FileInfo(photoPath);
                                    if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                                    { System.IO.File.Delete(photoPath); }

                                    pictureBox1.Load(photoPath);
                                }
                            }
                            catch
                            {
                                bmp = new Bitmap(Properties.Resources.LogoRYIK, pictureBox1.Width, pictureBox1.Height);
                                pictureBox1.BorderStyle = BorderStyle.None;
                                RefreshPictureBox(pictureBox1, bmp);
                            }
                        }
                    }
                }
                else if (sSelectedNAV.Contains("C")) //english С                       
                {
                    sTempNav = sSelectedNAV.Replace("C", "S");
                    try
                    {
                        pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                        photoPath = sFolderPhotos + sTempNav + @".jpg";

                        photoFileInfo = new System.IO.FileInfo(photoPath);
                        if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                        { System.IO.File.Delete(photoPath); }

                        pictureBox1.Load(photoPath);
                    }
                    catch
                    {
                        try
                        {
                            try { System.IO.Directory.CreateDirectory(sFolderPhotos); } catch { }
                            progressBar1.Value = 0;
                            using (System.Net.WebClient wc = new System.Net.WebClient())
                            {
                                wc.DownloadProgressChanged += wc_DownloadProgressChanged;
                                string url = "http://www.ais/usersimage/Fotos/" + sSelectedNAV + @".jpg";
                                string save_path = sFolderPhotos + sSelectedNAV + @".jpg";
                                wc.DownloadFileAsync(new Uri(url), save_path);
                                pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                                photoPath = sFolderPhotos + sSelectedNAV + @".jpg";

                                photoFileInfo = new System.IO.FileInfo(photoPath);
                                if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                                { System.IO.File.Delete(photoPath); }

                                pictureBox1.Load(photoPath);
                            }
                        }
                        catch
                        {
                            bmp = new Bitmap(Properties.Resources.LogoRYIK, pictureBox1.Width, pictureBox1.Height);
                            pictureBox1.BorderStyle = BorderStyle.None;
                            RefreshPictureBox(pictureBox1, bmp);
                        }
                    }
                }
                else
                {
                    try
                    {
                        try { System.IO.Directory.CreateDirectory(sFolderPhotos); } catch { }
                        progressBar1.Value = 0;
                        using (System.Net.WebClient wc = new System.Net.WebClient())
                        {
                            wc.DownloadProgressChanged += wc_DownloadProgressChanged;
                            string url = "http://www.ais/usersimage/Fotos/" + sSelectedNAV + @".jpg";
                            string save_path = sFolderPhotos + sSelectedNAV + @".jpg";
                            wc.DownloadFileAsync(new Uri(url), save_path);
                            pictureBox1.BorderStyle = BorderStyle.FixedSingle;
                            photoPath = sFolderPhotos + sSelectedNAV + @".jpg";

                            photoFileInfo = new System.IO.FileInfo(photoPath);
                            if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                            { System.IO.File.Delete(photoPath); }

                            pictureBox1.Load(photoPath);
                        }
                    }
                    catch
                    {
                        bmp = new Bitmap(Properties.Resources.LogoRYIK, pictureBox1.Width, pictureBox1.Height);
                        pictureBox1.BorderStyle = BorderStyle.None;
                        RefreshPictureBox(pictureBox1, bmp);
                    }
                }
            }
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

            try
            {
                if (s1 != null && s1.Length > 1)
                {
                    lString1 = new Label
                    {
                        Text = s1,
                        BackColor = Color.PaleGreen,
                        Location = new Point(this.Width - 245, 290),
                        Size = new Size(220, 20),
                        BorderStyle = BorderStyle.None,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Anchor = (AnchorStyles.Right | AnchorStyles.Top),
                        Parent = this
                    };
                    toolTip1.SetToolTip(lString1, "Номер владельца");
                }

                if (s2 != null && s2.Length > 1)
                {
                    lString2 = new Label
                    {
                        Text = s2,
                        BackColor = Color.PaleGreen,
                        Location = new Point(this.Width - 245, 320),
                        Size = new Size(220, 20),
                        BorderStyle = BorderStyle.None,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Anchor = (AnchorStyles.Right | AnchorStyles.Top),
                        Parent = this
                    };
                    toolTip1.SetToolTip(lString2, "ФИО владельца");
                }

                if (s3 != null && s3.Length > 1)
                {
                    lString3 = new Label
                    {
                        Text = s3,
                        BackColor = Color.PaleGreen,
                        Location = new Point(this.Width - 245, 350),
                        Size = new Size(220, 20),
                        BorderStyle = BorderStyle.None,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Anchor = (AnchorStyles.Right | AnchorStyles.Top),
                        Parent = this
                    };
                    toolTip1.SetToolTip(lString3, "Персональный код владельца");
                }

                if (s4 != null && s4.Length > 1)
                {
                    lString4 = new Label
                    {
                        Text = s4,
                        BackColor = Color.PaleGreen,
                        Location = new Point(this.Width - 245, 380),
                        Size = new Size(220, 20),
                        BorderStyle = BorderStyle.None,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Anchor = (AnchorStyles.Right | AnchorStyles.Top),
                        Parent = this
                    };
                    toolTip1.SetToolTip(lString4, "Организация, в которой работает владелец");
                }

                if (s5 != null && s5.Length > 1)
                {
                    lString5 = new Label
                    {
                        Text = s5,
                        BackColor = Color.PaleGreen,
                        Location = new Point(this.Width - 245, 410),
                        Size = new Size(220, 20),
                        BorderStyle = BorderStyle.None,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Anchor = (AnchorStyles.Right | AnchorStyles.Top),
                        Parent = this
                    };
                    toolTip1.SetToolTip(lString5, "Домашний адрес владельца");
                }

                if (s6 != null && s6.Length > 1)
                {
                    lString6 = new Label
                    {
                        Text = s6,
                        BackColor = Color.PaleGreen,
                        Location = new Point(this.Width - 245, 440),
                        Size = new Size(220, 20),
                        BorderStyle = BorderStyle.None,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Anchor = (AnchorStyles.Right | AnchorStyles.Top),
                        Parent = this
                    };
                    toolTip1.SetToolTip(lString6, "Дата увольнения");
                }

            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            try
            {
                photoFileInfo = new System.IO.FileInfo(photoPath);
                if (photoFileInfo.Exists && photoFileInfo.Length < 100)
                { System.IO.File.Delete(photoPath); }
            }
            catch { }
            photoFileInfo = null;
        }

        private void dataGridView1_DoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DefineSelectedDataInDatagrid();
            ShowData(sSelectedPhone, sSelectedFIO, sSelectedNAV, sSelected4, sSelected5, sSelected6);
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            DefineSelectedDataInDatagrid();
            ShowData(sSelectedPhone, sSelectedFIO, sSelectedNAV, sSelected4, sSelected5, sSelected6);
        }

        private void DefineSelectedDataInDatagrid() //Define Selected Data in Datagrid
        {
            sSelectedPhone = "";
            sSelectedFIO = "";
            sSelectedNAV = "";
            sSelected4 = "";
            sSelected5 = "";
            sSelected6 = "";

            lString1?.Dispose();
            lString2?.Dispose();
            lString3?.Dispose();
            lString4?.Dispose();
            lString5?.Dispose();
            lString6?.Dispose();
            if (dataGridView1.ColumnCount > 0)
            {
                string NameCollum = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].HeaderText.ToString(); //имя колонки выбранной ячейки

                int iIndexColumn1 = -1;   //dataGridView1.ColumnCount - всего колонок в датагрид отображается в данный момент     
                int iIndexColumn2 = -1;
                int iIndexColumn3 = -1;
                int iIndexColumn4 = -1;
                int iIndexColumn5 = -1;
                int iIndexColumn6 = -1;

                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "Номер телефона")
                    { iIndexColumn1 = i; }
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "ФИО")
                    { iIndexColumn2 = i; }
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV")
                    { iIndexColumn3 = i; }
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "Подразделение")
                    { iIndexColumn4 = i; }
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "Адрес")
                    { iIndexColumn5 = i; }
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "Дата увольнения")
                    { iIndexColumn6 = i; }
                }

                int selectedRowCount = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected);
                int IndexCurrentRow = dataGridView1.CurrentRow.Index;
                sSelectedPhone = dataGridView1.Rows[IndexCurrentRow].Cells[iIndexColumn1].Value.ToString();
                sSelectedFIO = dataGridView1.Rows[IndexCurrentRow].Cells[iIndexColumn2].Value.ToString();
                sSelectedNAV = dataGridView1.Rows[IndexCurrentRow].Cells[iIndexColumn3].Value.ToString();
                if (iIndexColumn4 > -1) sSelected4 = dataGridView1.Rows[IndexCurrentRow].Cells[iIndexColumn4].Value.ToString();
                if (iIndexColumn5 > -1) sSelected5 = dataGridView1.Rows[IndexCurrentRow].Cells[iIndexColumn5].Value.ToString();
                if (iIndexColumn6 > -1) sSelected6 = dataGridView1.Rows[IndexCurrentRow].Cells[iIndexColumn6].Value.ToString();
            }
        }

        private void FolderItem_Click(object sender, EventArgs e)
        { System.Diagnostics.Process.Start("explorer", Environment.CurrentDirectory); }

        private void HelpItem_Click(object sender, EventArgs e)
        { MakeHelp(); }

        private void MakeHelp()
        {
            System.Diagnostics.FileVersionInfo myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            String strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            DialogResult result = MessageBox.Show(
                "Справочник телефонов\n" + "Версия: " + myFileVersionInfo.FileVersion + "\nBuild: " +
                strVersion + "\n" + myFileVersionInfo.LegalCopyright +
                "\n\n1. Первый раз, перед получением данных с сервера T-Factura, необходимо:\n" + @"    A) Ввести адрес сервера в виде - SERVER.DOMAIN.SUBDOMAIN" + "\n" +
                @"    B) Ввести авторизационные данные" +
                "\n" + @"    С) Нажать кнопку " + "\"Сохранить\".\n" +
                "2. Корректный адрес сервера, имя и пароль пользователя T-Factura, можно получить в ИТ - отделе.\n" +
                "Для корректного импорта списка, с информацией о персонале в программу, используйте.\n\n" +
                "\nOriginal file: " + myFileVersionInfo.OriginalFilename + "\nFull path: " + Application.ExecutablePath,
                "Информация об использовании программы",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        private void textBoxData_Click(object sender, EventArgs e)
        { textBoxData.Clear(); }

        private void TableToTableshow(DataTable dt, DataTable dtShow)
        {
            dtShow.Clear();
            dtShow = dt.Clone();
            iRowRecords = 1;
            foreach (DataRow row in dt.Rows)
            {
                try
                {
                    dtShow.ImportRow(row);
                    iRowRecords++;
                }
                catch (Exception expt) { MessageBox.Show("dtShow\n" + expt.ToString()); }
            }
            ShowDataTableAtDatagrid(dtShow);
        }

        private void TableSearchToTableshow(DataTable dt, DataTable dtShow, string searchData)
        {
            dtShow.Rows.Clear();
            foreach (DataColumn column in dt.Columns)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row[column].ToString().ToLower().Contains(searchData.ToLower()))
                    {
                        try
                        { dtShow.ImportRow(row); }
                        catch { }
                    }
                }
            }
            iRowRecords = dtShow.Rows.Count;
            ShowDataTableAtDatagrid(dtShow);
        }

        private void ShowDataTableAtDatagrid(DataTable dt) //Access into Datagrid from other threads
        {
            try
            {
                if (this.InvokeRequired)
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        dataGridView1.DataSource = dt;
                        dataGridView1.AutoResizeColumns();
                    }));
                else
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.AutoResizeColumns();
                }
            }
            catch { }
        }

        private void TableSearchToLdata(DataTable dt, HashSet<string> hsData, string searchData)
        {
            iRowFIO = 0;
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Equals(searchData))
                    {
                        hsData.Add(row[column].ToString());
                        iRowFIO++;
                    }
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {            ApplicationExit();        }
    }
}
