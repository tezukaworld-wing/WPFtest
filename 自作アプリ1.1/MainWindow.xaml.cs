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
using Outlook = Microsoft.Office.Interop.Outlook;
using ClosedXML.Excel;

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
        public void syukin(object sender, EventArgs e)
        {
            var ol = new Outlook.Application();
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            DateTime dt = DateTime.Now;
            mail.Subject = "出勤:" + dt.ToString("yyyy年MM月dd日 ") + "(" + dt.ToString("ddd") + ")" + "手塚";
            Outlook.AddressEntry currentUser = ol.Session.CurrentUser.AddressEntry;
            mail.Body = "";
            mail.Recipients.Add("attendance@world-wing.com");
            mail.Recipients.ResolveAll();
            mail.Send();
        }
        private void nippou(object sender, EventArgs e)
        {
            var ol = new Outlook.Application();
            DateTime dt = DateTime.Now;
            String filePath = @"C:\Users\developer\Desktop\自作アプリ1.1\テキストテンプレート.xlsx";
            XLWorkbook workbook = new XLWorkbook(filePath);
            IXLWorksheet worksheet = workbook.Worksheet(1);
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = "［日報］手塚 " + dt.ToString("yyyy年MM月dd日 ") + "(" + dt.ToString("ddd") + ")";
            Outlook.AddressEntry currentUser = ol.Session.CurrentUser.AddressEntry;
            mail.Body = worksheet.Cell("A2").Value + "";
            mail.Recipients.Add("freshers-iar@world-wing.com");
            mail.Recipients.ResolveAll();
            mail.Send();
        }
        private void taikin(object sender, EventArgs e)
        {
            var ol = new Outlook.Application();
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            DateTime dt = DateTime.Now;
            mail.Subject = "退勤:" + dt.ToString("yyyy年MM月dd日 ") + "(" + dt.ToString("ddd") + ")" + "手塚";
            Outlook.AddressEntry currentUser = ol.Session.CurrentUser.AddressEntry;
            mail.Body = "本日の業務\r\n検証作業\r\nプログラミング";
            mail.Recipients.Add("attendance@world-wing.com");
            mail.Recipients.ResolveAll();
            mail.Send();
        }
        private void leport(object sender, EventArgs e)
        {
            var ol = new Outlook.Application();
            DateTime dt = DateTime.Now;
            String filePath = @"C:\Users\developer\Desktop\自作アプリ1.1\テキストテンプレート.xlsx";
            XLWorkbook workbook = new XLWorkbook(filePath);
            IXLWorksheet worksheet = workbook.Worksheet(1);
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = "［レポート］手塚 " + dt.ToString("yyyy年MM月dd日 ") + "(" + dt.ToString("ddd") + ")";
            Outlook.AddressEntry currentUser = ol.Session.CurrentUser.AddressEntry;
            mail.Body = worksheet.Cell("B2").Value + "";
            mail.Recipients.Add("manabiya-report@world-wing.com");
            mail.Recipients.ResolveAll();
            mail.Send();
        }
    }
}
