using System.Windows;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using Microsoft.Web.Administration;
using System.Diagnostics;
using System.Security.Principal;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.Net;

namespace Config_Util
{
    public partial class MainWindow
    {
        static string CMMSPathVar = string.Empty;
        bool GetServiceSuccess = false;
        string ServiceConfigPath = string.Empty;
        string RestConfigPath = string.Empty;
        string WebConfigPath = string.Empty;
        string ClientConfigPath = string.Empty;
        string ExePath = string.Empty;
        string ServicePath = string.Empty;
        string RestPath = string.Empty;
        string WebPath = string.Empty;
        string ClientPath = string.Empty;
        string ConnectionString = string.Empty;
        string SourcePath = string.Empty;
        string DestintantionPath = string.Empty;
        string ServiceTempPath = string.Empty;
        string NewestVersion = string.Empty;
        bool EditFlag = false;
        bool ComboAsSource = false;
        bool PathAsSource = false;

        string IsPredefined = System.Configuration.ConfigurationManager.AppSettings["IsPredefined"];
        public MainWindow()
        {
            InitializeComponent();
            /* CHECK FOR APP.CONFIG PREDEFINED PATHS */
            if (IsPredefined == "1")
            {
                CMMSPathCombo.Visibility = System.Windows.Visibility.Visible;
                ComboAsSource = true;
                /* GETS PREDEFINED PATHS FROM APP.CONFIG AND POPULATES COMBOBOX WITH THEM */
                foreach (string key in System.Configuration.ConfigurationManager.AppSettings.AllKeys)
                {
                    if (key.StartsWith("CMMS"))
                    {
                        string value = System.Configuration.ConfigurationManager.AppSettings[key];
                        CMMSPathCombo.Items.Add(value);
                    }
                }
                /* */
                if (CMMSPathCombo.Items.Count < 0)
                {
                    MessageBox.Show("No directory to choose. Check app.config.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    System.Environment.Exit(0);
                }
                else
                {
                    CMMSPathCombo.SelectedIndex = 0;
                }
            }
            else
            {
                CMMSPath.Visibility = System.Windows.Visibility.Visible;
                PathAsSource = true;
            }

        }
        bool IsElevated
        {
            get
            {
                return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        public void LoadFilesButton_Click(object sender, RoutedEventArgs e)

        {
            bool GetServiceFail = false;
            string CMMSPathVar = string.Empty;
            LoadFilesButton.Content = "Loading...";
            if (PathAsSource == true)
            {
                CMMSPathVar = CMMSPath.Text;
            }
            else if (ComboAsSource == true)
            {
                CMMSPathVar = CMMSPathCombo.SelectedItem.ToString();

            }
            if (string.IsNullOrEmpty(CMMSPathVar))
            {
                LoadFilesButton.Content = "Load CMMS Catalogue";
            }
            else
            {
                if (Directory.Exists(CMMSPathVar))
                {
                    string NewServicePath = CMMSPathVar + "\\Service";
                    string OldServicePath = CMMSPathVar + "\\RRM3Services";
                    string NewRestPath = CMMSPathVar + "\\RestService";
                    string OldRestPath = CMMSPathVar + "\\RRM3RestService";
                    string NewWebPath = CMMSPathVar + "\\WebClient";
                    string OldWebPath = CMMSPathVar + "\\RRM3WebClient";
                    string NewClientPath = CMMSPathVar + "\\Client";
                    string OldClientPath = CMMSPathVar + "\\RRM3Client";
                    string NewExePath = CMMSPathVar + "\\Client\\RRM3.exe";
                    string OldExePath = CMMSPathVar + "\\RRM3Client\\RRM3.exe";
                    string PanelPath = CMMSPathVar + "\\Panel";

                    if (Directory.Exists(NewServicePath))
                    {
                        ServicePath = NewServicePath;
                        ServiceBox.Text = NewServicePath;
                        ServiceConfigPath = NewServicePath + "\\Web.config";
                        ServiceTempPath = NewServicePath + "\\Temp";
                        GetServiceSuccess = true;
                    }
                    else if (Directory.Exists(OldServicePath))
                    {
                        ServicePath = OldServicePath;
                        ServiceBox.Text = OldServicePath;
                        ServiceConfigPath = OldServicePath + "\\Web.config";
                        ServiceTempPath = OldServicePath + "\\Temp";
                        GetServiceSuccess = true;
                    }
                    else
                    {
                        MessageBox.Show("Service directory not found! Check if CMMS path is correct.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                        LoadFilesButton.Content = "Load CMMS Catalogue";

                    }
                    if (GetServiceSuccess == true)
                    {
                        if (Directory.Exists(NewRestPath))
                        {
                            RestPath = NewRestPath;
                            RestBox.Text = NewRestPath;
                            RestConfigPath = NewRestPath + "\\Web.config";
                        }
                        else if (Directory.Exists(OldRestPath))
                        {
                            RestPath = OldRestPath;
                            RestBox.Text = OldRestPath;
                            RestConfigPath = OldRestPath + "\\Web.config";
                        }
                        else
                        {
                            RestBox.Text = "Rest not found";
                        }
                        if (Directory.Exists(NewWebPath))
                        {
                            WebPath = NewWebPath;
                            WebBox.Text = NewWebPath;
                            WebConfigPath = NewWebPath + "\\Web.config";
                        }
                        else if (Directory.Exists(OldWebPath))
                        {
                            WebPath = OldWebPath;
                            WebBox.Text = OldWebPath;
                            WebConfigPath = OldWebPath + "\\Web.config";
                        }
                        else
                        {
                            WebBox.Text = "Web not found";
                        }
                        if (Directory.Exists(NewClientPath))
                        {
                            ClientBox.Text = NewClientPath;
                            ClientConfigPath = NewClientPath + "\\RRM3.exe.config";
                            ExePath = NewClientPath + "\\RRM3.exe";
                        }
                        else if (Directory.Exists(OldClientPath))
                        {
                            ClientBox.Text = OldClientPath;
                            ClientConfigPath = OldClientPath + "\\RRM3.exe.config";
                            ExePath = OldClientPath + "\\RRM3.exe";
                        }
                        else
                        {
                            ClientBox.Text = "Client not found";
                        }
                        if (Directory.Exists(PanelPath))
                        {
                            PanelBox.Text = PanelPath;
                        }
                        else
                        {
                            PanelBox.Text = "Panel not found.";
                        }
                        XmlDocument config = new XmlDocument();
                        XmlDocument clientconfig = new XmlDocument();
                        clientconfig.Load(ClientConfigPath);
                        config.Load(ServiceConfigPath);
                        XmlNode client_node = clientconfig.DocumentElement.SelectSingleNode("/configuration/appSettings/add");
                        XmlNode node = config.DocumentElement.SelectSingleNode("/configuration/connectionStrings/add");
                        string ClientVersion = client_node.Attributes["value"].Value;
                        string ConnectionString = node.Attributes["connectionString"].Value;
                        Regex regex = new Regex("=(\\w.*?);");
                        MatchCollection result = regex.Matches(ConnectionString);
                        AddressBox.Text = result[2].ToString();
                        AddressBox.Text = AddressBox.Text.Substring(1, AddressBox.Text.Length - 2);
                        NameBox.Text = result[3].ToString();
                        NameBox.Text = NameBox.Text.Substring(1, NameBox.Text.Length - 2);
                        LoginBox.Text = result[5].ToString();
                        LoginBox.Text = LoginBox.Text.Substring(1, LoginBox.Text.Length - 2);
                        PasswordBox.Text = result[6].ToString();
                        PasswordBox.Text = PasswordBox.Text.Substring(1, PasswordBox.Text.Length - 2);
                        VersionBox.Text = ClientVersion;
                        IISDataButton.IsEnabled = true;
                        EditButton.IsEnabled = true;
                        LoadDBButton.IsEnabled = true;
                        RunButton.IsEnabled = true;
                        CheckUpdateButton.IsEnabled = true;
                        LoadFilesButton.Content = "Reload CMMS Catalogue";
                    }
                }
                else if (GetServiceFail == false)
                {
                    MessageBox.Show("That is not valid CMMS directory.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    LoadFilesButton.Content = "Load CMMS Catalogue";
                }

            }
        }



        private void IISDataButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IISTab.IsEnabled = true;
                IISTab.Focus();
                ServerManager ServerMgr = new ServerManager();
                foreach (var site in ServerMgr.Sites)
                {
                    if (site.Applications["/"].VirtualDirectories["/"].PhysicalPath != null || site.Applications["/"].VirtualDirectories["/"].PhysicalPath != string.Empty)
                    {
                        if (site.Applications["/"].VirtualDirectories["/"].PhysicalPath == ServicePath)
                        {
                            ServiceSiteBox.Text = site.Name;
                            foreach (var binding in site.Bindings)
                            {
                                if (binding.Protocol == "net.tcp")
                                {
                                    string ConvertedBinding = binding.BindingInformation;
                                    Regex regex = new Regex("\\d+");
                                    Match service_result = regex.Match(ConvertedBinding);
                                    ServicePortBox_NetTcp.Text = service_result.Value;
                                }
                                else
                                {
                                    string ConvertedBinding = binding.BindingInformation;
                                    Regex regex = new Regex("\\d+");
                                    Match service_result = regex.Match(ConvertedBinding);
                                    ServicePortBox_Http.Text = service_result.Value;
                                }
                            }
                        }
                        if (site.Applications["/"].VirtualDirectories["/"].PhysicalPath == RestPath)
                        {
                            RestSiteBox.Text = site.Name;
                            foreach (var binding in site.Bindings)
                            {
                                RestPortBox.Visibility = System.Windows.Visibility.Visible;
                                string ConvertedBinding = binding.BindingInformation;
                                Regex regex = new Regex("\\d+");
                                Match rest_result = regex.Match(ConvertedBinding);
                                RestPortBox.Text = rest_result.Value;
                            }
                        }
                        if (site.Applications["/"].VirtualDirectories["/"].PhysicalPath == WebPath)
                        {
                            WebSiteBox.Text = site.Name;
                            foreach (var binding in site.Bindings)
                            {
                                WebPortBox.Visibility = System.Windows.Visibility.Visible;
                                string ConvertedBinding = binding.BindingInformation;
                                Regex regex = new Regex("\\d+");
                                Match web_result = regex.Match(ConvertedBinding);
                                WebPortBox.Text = web_result.Value;
                            }
                        }
                    }
                    else
                    {

                    }
                }
            }

            catch
            {
                MessageBox.Show("App requires elevation to access IIS Services.", "Run as administrator", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    

    
        public void EditButton_Click(object sender, RoutedEventArgs e)
        {
            /* BUTTON NOT IN SAVE MODE */
                if (EditFlag == false)
            {
                /* SET FOCUS, UNLOCK BOXES, LOCK APP - SET SAVE MODE */
                EditFlag = true;
                EditButton.Content = "Save files";
                IISTab.IsEnabled = false;
                FilesTab.IsEnabled = false;
                BasicTab.Focus();
                AddressBox.IsReadOnly = false;
                NameBox.IsReadOnly = false;
                LoginBox.IsReadOnly = false;
                PasswordBox.IsReadOnly = false;
            }
            /* BUTTON IN SAVE MODE */
            else if (EditFlag == true)
            {
                
            }
            else
            {
                MessageBox.Show("You weren't supposed to see this error, ever.", "I'm not even mad", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = ExePath;
            Process.Start(start);
        }
        private void label_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Clipboard.SetText("Admin" + VersionBox.Text);
        }
        /* DECLARATION OF NEW SQL CONNECTION ISNTANCE */
        SqlConnection Connection = new SqlConnection();
        private void LoadDBButton_Click(object sender, RoutedEventArgs e)
        {
            string DBModuleConnectionString = "Server=" + AddressBox.Text + ";Database=" + NameBox.Text + ";User ID=" + LoginBox.Text + ";Password=" + PasswordBox.Text + ";";
            Connection.ConnectionString = DBModuleConnectionString;
            try
            {
                Connection.Open();
            }
            catch
            {
                MessageBox.Show("Could not establish connection to a DataBase", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            string ConnectionStatus = Connection.State.ToString();
            SQLTab.IsEnabled = true;
            SQLTab.Focus();
            DBAddressLabel.Content = AddressBox.Text;
            DBNameLabel.Content = NameBox.Text;
            EditButton.IsEnabled = false;
            CMMSPath.IsEnabled = false;
            CMMSPathCombo.IsEnabled = false;
        }
        private void LoadSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            string LoadSettingsCommand = "SELECT [Nazwa], [Wartosc] FROM dbo.UstawieniaAplikacji WHERE Id in (1,6,17,23,25,46,53)";
            SqlCommand LoadSettingsQuery = new SqlCommand(LoadSettingsCommand, Connection);
            SqlDataAdapter SettingsDataAdapter = new SqlDataAdapter();
            DataSet SettingsDataSet = new DataSet();
            SettingsDataAdapter.SelectCommand = LoadSettingsQuery;
            SettingsDataAdapter.Fill(SettingsDataSet);
            for (i = 0; i <= SettingsDataSet.Tables[0].Rows.Count - 1; i++)
            {
               
            }
        }
        private void SetSettingsButton_Click(object sender, RoutedEventArgs e)
        {

        }
        private void ClearBinButton_Click(object sender, RoutedEventArgs e)
        {

        }
        private void CloseConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            Connection.Close();
            BasicTab.Focus();
            SQLTab.IsEnabled = false;
            CMMSPath.IsEnabled = true;
            CMMSPathCombo.IsEnabled = true;
        }
        private void CheckUpdateButton_Click(object sender, RoutedEventArgs e)
        {
            /* sprawdza dostępność wersji na serwerze queris 57 */
            string RCDir = "\\\\queris57\\Builds\\RRM3 Release Candidate";
            List<string> Directories = new List<string>(Directory.EnumerateDirectories(RCDir));
            Regex DirRegex = new Regex("RRM3_.+");
            for (int i = 0; i < Directories.Count; i++)
            {
                string DirToRegex = Directories[i];
                Match DirRegexMatch = DirRegex.Match(DirToRegex);
                Directories[i] = DirRegexMatch.ToString();
            }
            NewestVersion = Directories[(Directories.Count-1)];
            if (NewestVersion != VersionBox.Text)
            {
                NewVersionLabel.Text = "Version " + NewestVersion + " is available for update.";
                if (Directory.Exists(ServiceTempPath + "\\" + NewestVersion))
                {
                    UpdateButton.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("There is new version available on queris57. Please download manually.", "Update has been canceled.", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            else if (NewestVersion == VersionBox.Text)
            {
                NewVersionLabel.Text = "This version is up to date";
                UpdateButton.IsEnabled = false;
            }
            else
            {
                UpdateButton.IsEnabled = false;
                NewVersionLabel.Text = "System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. System Failure. ";
            }
        }
        private void UpdateButton_Click(object sender, RoutedEventArgs e)
        {

            if (MessageBox.Show("Are you sure you want to update to newer RC version?", "Update", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                foreach (string newPath in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
                    File.Copy(newPath, newPath.Replace(SourcePath, DestinationPath), true);
            }
            else
            {
                MessageBox.Show("Update has been cancelled by user.", "Update has been canceled.", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

    }
}
