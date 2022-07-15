using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Threading;


namespace 自作アプリ1号
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {

        }

        public void Zyobukan(object sender, EventArgs e)
        {
            ProcessStartInfo pi = new ProcessStartInfo()
            {
                FileName = "https://id.jobcan.jp/users/sign_in",
                UseShellExecute = true,
            };
            Process.Start(pi);
        }

        public void hyouka(object sender, EventArgs e)
        {
            ProcessStartInfo pi = new ProcessStartInfo()
            {
                FileName = "https://worldwing.sharepoint.com/:x:/r/sites/manabiya/_layouts/15/Doc.aspx?sourcedoc=%7BF9FDCA0D-1EA8-4950-89A7-C582A5A63DC0%7D&file=schoo%E8%A6%96%E8%81%B4%E5%B1%A5%E6%AD%B4.xlsx&action=default&mobileredirect=true&wdLOR=c36AF46A4-562F-4F3F-BCA9-9CEF4F7E55A4&cid=0704a3a7-b45d-407d-94f0-a495cc2a7870",
                UseShellExecute = true,
            };
            Process.Start(pi);
        }


    }
}
