using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
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
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

        public void LoadFilesButton_Click(object sender, RoutedEventArgs e)
        {
            string CMMSPathVar = CMMSPath.Text;
            if (string.IsNullOrEmpty(CMMSPathVar))
            {
                MessageBox.Show("Enter CMMS Path first!", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                if (Directory.Exists(CMMSPathVar))
                {
                    string ServiceConfigPath = string.Empty;
                    string RestConfigPath = string.Empty;
                    string WebConfigPath = string.Empty;
                    string ClientConfigPath = string.Empty;
                    string ExePath = string.Empty;
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

                    if (Directory.Exists(NewServicePath))
                    {
                        ServiceConfigPath = NewServicePath + "Web.config";
                    }
                    else if (Directory.Exists(OldServicePath))
                    {
                        ServiceConfigPath = OldServicePath + "Web.config";
                    }
                    else
                    {
                        MessageBox.Show("Service directory not found! Check if CMMS path is correct.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    if (Directory.Exists(NewRestPath))
                    {
                        RestConfigPath = NewRestPath + "Web.config";
                    }
                    else if (Directory.Exists(OldRestPath))
                    {
                        RestConfigPath = OldRestPath + "Web.config";
                    }
                    else
                    {
                        MessageBox.Show("Rest Service directory not found! Check if CMMS path is correct.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    if (Directory.Exists(NewWebPath))
                    {
                        WebConfigPath = NewWebPath + "Web.config";
                    }
                    else if (Directory.Exists(OldWebPath))
                    {
                        WebConfigPath = OldWebPath + "Web.config";
                    }
                    else
                    {
                        MessageBox.Show("Web Service directory not found! Check if CMMS path is correct.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    if (Directory.Exists(NewClientPath))
                    {
                        ClientConfigPath = NewClientPath + "Web.config";
                        ExePath = NewClientPath + "RRM3.exe";
                    }
                    else if (Directory.Exists(OldClientPath))
                    {
                        ClientConfigPath = OldClientPath + "Web.config";
                        ExePath = OldClientPath + "RRM3.exe";
                    }
                    else
                    {
                        MessageBox.Show("Client directory not found! Check if CMMS path is correct.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    XmlDocument config = new XmlDocument();
                    config.Load(ServiceConfigPath);
                    XmlNode node = config.DocumentElement.SelectSingleNode("/configuration/connectionStrings/add");
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
                    AddressBox.Visibility = System.Windows.Visibility.Visible;
                    NameBox.Visibility = System.Windows.Visibility.Visible;
                    LoginBox.Visibility = System.Windows.Visibility.Visible;
                    PasswordBox.Visibility = System.Windows.Visibility.Visible;
                    VersionBox.Visibility = System.Windows.Visibility.Visible;
                    AddressLabel.Visibility = System.Windows.Visibility.Visible;
                    NameLabel.Visibility = System.Windows.Visibility.Visible;
                    LoginLabel.Visibility = System.Windows.Visibility.Visible;
                    PasswordLabel.Visibility = System.Windows.Visibility.Visible;
                    VersionLabel.Visibility = System.Windows.Visibility.Visible;
                    ListFilesButton.IsEnabled = true;
                    EditButton.IsEnabled = true;
                    LoadDBButton.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("Directory not found. Make sure path is correct.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
            }
        }

    }
}
