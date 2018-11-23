using System.Collections.Generic;
using System.Windows;
using ExcelOperator;
using Microsoft.Win32;
using TenderingExpert.Data;
using WordOperator;

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

        public TenderForm TenderExcelForm { get; set; } = new TenderForm();

        public List<PackageInformation> PackageInformations { get; set; }

        private WordReader tenderWordReader;

        private WordReader purchaseWordReader;

        private ExcelWriter excelWriter;

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
            var tenderDocText = TenderDoc.Text;
            if (!string.IsNullOrEmpty(tenderDocText))
            {
                tenderWordReader = new WordReader(tenderDocText);
                tenderWordReader.StartRead();

                TenderInformation.LoadInfo(tenderWordReader);
                PackageInformations = TenderInformation.LoadPackageInfo(tenderWordReader);
            }

            var purchaseDocText = PurchaseDoc.Text;
            if (!string.IsNullOrEmpty(purchaseDocText))
            {
                purchaseWordReader = new WordReader(purchaseDocText);
                purchaseWordReader.StartRead();

                for (int i = 0; i < PackageInformations.Count; i++)
                {
                    PackageInformations[i].PurchaseInformations =
                        TenderInformation.LoadPurchaseInfo(purchaseWordReader, i + 1);
                }
            }

            PackageList.ItemsSource = PackageInformations;
        }

        private void Window_Closed(object sender, System.EventArgs e)
        {
            tenderWordReader?.Close();
            purchaseWordReader?.Close();
            excelWriter?.Close();
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
            excelWriter = new ExcelWriter();
            excelWriter.Create();
            
            TenderExcelForm.TenderInfo = TenderInformation;
            TenderExcelForm.PackageInfo = PackageInformations[0];
            
            TenderExcelForm.Init();
            TenderExcelForm.FillContent(excelWriter);
            
            excelWriter.SaveAs("D:\\评标表.xls");
            excelWriter.Close();
        }
    }
}
