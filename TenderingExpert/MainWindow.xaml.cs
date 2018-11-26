using System;
using System.Collections.Generic;
using System.ComponentModel;
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

        private string tenderDocText;

        private string purchaseDocText;

        private string excelPath;

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

            StartRead.IsEnabled = false;
            StartRead.Content = "正在加载";

            tenderDocText = TenderDoc.Text;
            purchaseDocText = PurchaseDoc.Text;

            StartBackgroundWork(LoadWordInfo,null,LoadRunComplete);
        }

        private void StartBackgroundWork(DoWorkEventHandler doWork, ProgressChangedEventHandler progressChanged, RunWorkerCompletedEventHandler completed)
        {
            var work = new BackgroundWorker {WorkerReportsProgress = true};
            work.DoWork += doWork;
            work.ProgressChanged += progressChanged;
            work.RunWorkerCompleted += completed;
            work.RunWorkerAsync();
        }

        private void LoadWordInfo(object sender, DoWorkEventArgs e)
        {
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
                    Dispatcher.Invoke(() => { ShowErrorMessage(exception.Message); });
                }
            }

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
                    Dispatcher.Invoke(() => { ShowErrorMessage(exception.Message); });
                }
            }
        }

        private void CreateExcels(object sender, DoWorkEventArgs e)
        {
            foreach (PackageInformation information in PackageInformations)
            {
                try
                {
                    var excelWriter = new ExcelWriter();
                    excelWriter.Create();

                    var tenderExcelForm = new TenderForm { TenderInfo = TenderInformation, PackageInfo = information };

                    tenderExcelForm.Init(PackageInformations.Count == 1);
                    tenderExcelForm.FillContent(excelWriter);

                    var path = Path.Combine(excelPath, information.DeviceName);
                    path = path.Replace("/", "");
                    excelWriter.SaveAs(path);
                    excelWriter.Close();
                }
                catch (Exception exception)
                {
                    ShowErrorMessage(exception.Message);
                }
            }
        }

        private void LoadRunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                PackageList.ItemsSource = PackageInformations;
                ShowInfoMessage("加载完成");
                StartRead.IsEnabled = true;
                StartRead.Content = "读取";

            });
        }

        private void CreateRunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                ShowInfoMessage("创建完成");
                CreateExcel.IsEnabled = true;
                CreateExcel.Content = "创建";
            });
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

            CreateExcel.IsEnabled = false;
            CreateExcel.Content = "正在创建";

            excelPath = ExcelPath.Text;

            StartBackgroundWork(CreateExcels, null, CreateRunComplete);
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
