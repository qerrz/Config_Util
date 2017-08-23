using System.Windows;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using Microsoft.Web.Administration;
using System.Diagnostics;
using System.Security.Principal;
using System.Threading;

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
        //string CMMSTestingPath = "C:\\Queris\\CMMS - Testing";
        bool EditFlag = false;
        bool ComboAsSource = false;
        bool PathAsSource = false;
        string IsPredefined = System.Configuration.ConfigurationManager.AppSettings["IsPredefined"];
        public MainWindow()
        {
            InitializeComponent();
            if (IsPredefined == "1")
            {
                CMMSPathCombo.Visibility = System.Windows.Visibility.Visible;
                ComboAsSource = true;
                foreach (string key in System.Configuration.ConfigurationManager.AppSettings.AllKeys)
                {
                    if (key.StartsWith("CMMS"))
                    {
                        string value = System.Configuration.ConfigurationManager.AppSettings[key];
                        CMMSPathCombo.Items.Add(value);
                    }
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
            // CMMSPath.Text = "C:\\Queris\\CMMS - Testing";
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
                MessageBox.Show("Enter CMMS Path first!", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                LoadFilesButton.Content = "Load CMMS Catalogue";
                //CMMSPath.Text = CMMSTestingPath;
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
                        GetServiceSuccess = true;
                    }
                    else if (Directory.Exists(OldServicePath))
                    {
                        ServicePath = OldServicePath;
                        ServiceBox.Text = OldServicePath;
                        ServiceConfigPath = OldServicePath + "\\Web.config";
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
                        //LoadDBButton.IsEnabled = true;
                        RunButton.IsEnabled = true;
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
            if (IsElevated == true)
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
            else
            {
                MessageBox.Show("You need to run this app as administrator to use this option.", "Run as administrator", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                string ConnectionString = string.Empty;
                /* GENERATE NEW CONNECTIONSTRING */
                ConnectionString = "metadata=res://*/RrmDBModel.csdl|res://*/RrmDBModel.ssdl|res://*/RrmDBModel.msl;provider=System.Data.SqlClient;provider connection string=\"data source=" + AddressBox.Text + ";initial catalog=" + NameBox.Text + ";persist security info=True;user id=" + LoginBox.Text + ";password=" + PasswordBox.Text + ";MultipleActiveResultSets=True;App=EntityFramework\"";
                /* BEGIN SAVING DOCUMENTS */
                XmlDocument ServiceFile = new XmlDocument();
                XmlDocument RestFile = new XmlDocument();
                XmlDocument WebFile = new XmlDocument();
                /* LOAD DOCUMENTS */
                ServiceFile.Load(ServiceConfigPath);
                RestFile.Load(RestConfigPath);
                WebFile.Load(WebConfigPath);
                /* SELECT NODES */
                XmlNode ServiceConnStringNode = ServiceFile.SelectSingleNode("/configuration/connectionStrings/add");
                XmlNode RestFileNode = RestFile.SelectSingleNode("/configuration/connectionStrings/add");
                XmlNode WebFileNode = WebFile.SelectSingleNode("/configuration/connectionStrings/add");
                /* CHANGE NODES */
                ServiceConnStringNode.Attributes["connectionString"].Value = ConnectionString;
                RestFileNode.Attributes["connectionString"].Value = ConnectionString;
                WebFileNode.Attributes["connectionString"].Value = ConnectionString;
                /* SAVE FILES */
                ServiceFile.Save(ServiceConfigPath);
                RestFile.Save(RestConfigPath);
                WebFile.Save(WebConfigPath);
                /* LOCK TEXTBOXES */
                AddressBox.IsReadOnly = true;
                NameBox.IsReadOnly = true;
                LoginBox.IsReadOnly = true;
                PasswordBox.IsReadOnly = true;
                /* UNLOCK APP - END SAVE MODE */
                EditFlag = false;
                EditButton.Content = "Edit Connection Data";
                IISTab.IsEnabled = true;
                FilesTab.IsEnabled = true;
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
    }
}
