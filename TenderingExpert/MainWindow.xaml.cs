using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using ExcelOperator;
using TenderingExpert.Data;
using WordOperator;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace TenderingExpert
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = TenderInformation;
        }

        public TenderingInformation TenderInformation { get; set; } = new TenderingInformation();

        public List<PackageInformation> PackageInformations { get; set; }

        private WordReader tenderWordReader;

        private WordReader purchaseWordReader;        

        private void SelectTenderFile_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                Filter = "Word File(*.doc, *.docx) |*doc;*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
                TenderDoc.Text = openFileDialog.FileName;
        }

        private void StartRead_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TenderDoc.Text) || string.IsNullOrEmpty(PurchaseDoc.Text))
            {
                ShowInfoMessage("请选择Word文件");
                return;
            }

            var tenderDocText = TenderDoc.Text;
            if (!string.IsNullOrEmpty(tenderDocText))
            {
                try
                {
                    tenderWordReader = new WordReader(tenderDocText);
                    tenderWordReader.StartRead();

                    TenderInformation.LoadInfo(tenderWordReader);
                    PackageInformations = TenderInformation.LoadPackageInfo(tenderWordReader);

                    tenderWordReader.Close();
                }
                catch (Exception exception)
                {
                    ShowErrorMessage(exception.Message);
                }
            }

            var purchaseDocText = PurchaseDoc.Text;
            if (!string.IsNullOrEmpty(purchaseDocText))
            {
                try
                {
                    purchaseWordReader = new WordReader(purchaseDocText);
                    purchaseWordReader.StartRead();

                    for (int i = 0; i < PackageInformations.Count; i++)
                    {
                        PackageInformations[i].PurchaseInformations =
                            TenderInformation.LoadPurchaseInfo(purchaseWordReader, i + 1);
                    }

                    purchaseWordReader.Close();
                }
                catch (Exception exception)
                {
                    ShowErrorMessage(exception.Message);
                }
            }

            PackageList.ItemsSource = PackageInformations;
        }

        private void Window_Closed(object sender, System.EventArgs e)
        {
            tenderWordReader?.Close();
            purchaseWordReader?.Close();

            tenderWordReader = null;
            purchaseWordReader = null;
        }

        private void SelectPurchaseFile_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                Filter = "Word File(*.doc, *.docx) |*doc;*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
                PurchaseDoc.Text = openFileDialog.FileName;
        }

        private void CreateExcel_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(ExcelPath.Text))
            {
                ShowInfoMessage("请选择Excel生成路径");
                return;
            }

            foreach (PackageInformation information in PackageInformations)
            {
                try
                {
                    var excelWriter = new ExcelWriter();
                    excelWriter.Create();

                    var tenderExcelForm = new TenderForm {TenderInfo = TenderInformation, PackageInfo = information};

                    tenderExcelForm.Init();
                    tenderExcelForm.FillContent(excelWriter);

                    var path = Path.Combine(ExcelPath.Text, information.DeviceName);
                    excelWriter.SaveAs(path);
                    excelWriter.Close();
                }
                catch (Exception exception)
                {
                    ShowErrorMessage(exception.Message);
                }
            }
        }

        private void SelectExcelPath_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new FolderBrowserDialog();

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ExcelPath.Text = openFileDialog.SelectedPath;
        }

        private void ShowInfoMessage(string msg)
        {
            MessageBox.Show(msg, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowErrorMessage(string msg)
        {
            MessageBox.Show(msg, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
