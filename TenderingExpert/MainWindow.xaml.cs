using System.Collections.Generic;
using System.Windows;
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
    }
}
