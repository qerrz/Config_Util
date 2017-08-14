using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using Microsoft.Web.Administration;
using System.Collections.Specialized;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Config_Util
{
    public partial class MainWindow
    {
        public void LoadFilesButton_Click(object sender, RoutedEventArgs e)
        {
            bool GetServiceFail = false;
            LoadFilesButton.Content = "Loading...";
            string CMMSPathVar = CMMSPath.Text;
            if (string.IsNullOrEmpty(CMMSPathVar))
            {
                MessageBox.Show("Enter CMMS Path first!", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                LoadFilesButton.Content = "Load CMMS Catalogue";
            }
            else
            {
                if (Directory.Exists(CMMSPathVar))
                {
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
                    string NewServicePath = CMMSPathVar + "\\Service\\";
                    string OldServicePath = CMMSPathVar + "\\RRM3Services\\";
                    string NewRestPath = CMMSPathVar + "\\RestService\\";
                    string OldRestPath = CMMSPathVar + "\\RRM3RestService\\";
                    string NewWebPath = CMMSPathVar + "\\WebClient\\";
                    string OldWebPath = CMMSPathVar + "\\RRM3WebClient\\";
                    string NewClientPath = CMMSPathVar + "\\Client\\";
                    string OldClientPath = CMMSPathVar + "\\RRM3Client\\";
                    string NewExePath = CMMSPathVar + "\\Client\\RRM3.exe";
                    string OldExePath = CMMSPathVar + "\\RRM3Client\\RRM3.exe";
                    string PanelPath = CMMSPathVar + "\\Panel\\";

                    if (Directory.Exists(NewServicePath))
                    {
                        ServiceBox.Text = NewServicePath;
                        ServiceConfigPath = NewServicePath + "Web.config";
                        GetServiceSuccess = true;
                    }
                    else if (Directory.Exists(OldServicePath))
                    {
                        ServiceBox.Text = OldServicePath;
                        ServiceConfigPath = OldServicePath + "Web.config";
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
                            RestBox.Text = NewRestPath;
                            RestConfigPath = NewRestPath + "Web.config";
                        }
                        else if (Directory.Exists(OldRestPath))
                        {
                            RestBox.Text = NewRestPath;
                            RestConfigPath = OldRestPath + "Web.config";
                        }
                        else
                        {
                            RestBox.Text = "Rest not found";
                        }
                        if (Directory.Exists(NewWebPath))
                        {
                            WebBox.Text = NewWebPath;
                            WebConfigPath = NewWebPath + "Web.config";
                        }
                        else if (Directory.Exists(OldWebPath))
                        {
                            WebBox.Text = OldWebPath;
                            WebConfigPath = OldWebPath + "Web.config";
                        }
                        else
                        {
                            WebBox.Text = "Web not found";
                        }
                        if (Directory.Exists(NewClientPath))
                        {
                            ClientBox.Text = NewClientPath;
                            ClientConfigPath = NewClientPath + "RRM3.exe.config";
                            ExePath = NewClientPath + "RRM3.exe";
                        }
                        else if (Directory.Exists(OldClientPath))
                        {
                            ClientBox.Text = OldClientPath;
                            ClientConfigPath = OldClientPath + "RRM3.exe.config";
                            ExePath = OldClientPath + "RRM3.exe";
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
                        AddressBox.Visibility = System.Windows.Visibility.Visible;
                        NameBox.Visibility = System.Windows.Visibility.Visible;
                        LoginBox.Visibility = System.Windows.Visibility.Visible;
                        PasswordBox.Visibility = System.Windows.Visibility.Visible;
                        VersionBox.Visibility = System.Windows.Visibility.Visible;
                        ServiceBox.Visibility = System.Windows.Visibility.Visible;
                        ClientBox.Visibility = System.Windows.Visibility.Visible;
                        RestBox.Visibility = System.Windows.Visibility.Visible;
                        WebBox.Visibility = System.Windows.Visibility.Visible;
                        PanelBox.Visibility = System.Windows.Visibility.Visible;
                        AddressLabel.Visibility = System.Windows.Visibility.Visible;
                        NameLabel.Visibility = System.Windows.Visibility.Visible;
                        LoginLabel.Visibility = System.Windows.Visibility.Visible;
                        PasswordLabel.Visibility = System.Windows.Visibility.Visible;
                        VersionLabel.Visibility = System.Windows.Visibility.Visible;
                        ServiceLabel.Visibility = System.Windows.Visibility.Visible;
                        ClientLabel.Visibility = System.Windows.Visibility.Visible;
                        RestLabel.Visibility = System.Windows.Visibility.Visible;
                        WebLabel.Visibility = System.Windows.Visibility.Visible;
                        PanelLabel.Visibility = System.Windows.Visibility.Visible;
                        IISDataButton.IsEnabled = true;
                        EditButton.IsEnabled = true;
                        LoadDBButton.IsEnabled = true;
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
            ServerManager ServerMgr = new ServerManager();
            foreach (var site in ServerMgr.Sites)
               /* if (site.IsLocallyStored == ServicePath) */
                {

                }
        }
    }
}
