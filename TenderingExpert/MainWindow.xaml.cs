using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
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

        private int currentPage = 1;

        #region ButtonClick

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

            StartBackgroundWork(LoadWordInfo, null, LoadRunComplete);
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

        private void PrePage_Click(object sender, RoutedEventArgs e)
        {
            if (currentPage > 1)
            {
                currentPage -= 1;
                UpdateContent();
            }
        }

        private void NextPage_Click(object sender, RoutedEventArgs e)
        {
            if (currentPage < GetContentPageCount())
            {
                currentPage += 1;
                UpdateContent();
            }
        }

        private void JumpToPage_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Page.Text))
                return;

            int page = Convert.ToInt32(Page.Text);
            currentPage = page;
            UpdateContent();
        }

        private void AutoRead_Click(object sender, RoutedEventArgs e)
        {
            if (tenderWordReader == null || purchaseWordReader == null)
            {
                ShowInfoMessage("未读取任何内容");
                return;
            }

            AutoRead.IsEnabled = false;
            AutoRead.Content = "正在读取";

            StartBackgroundWork((o, args) =>
            {
                TenderInformation.LoadInfo(tenderWordReader);
                PackageInformations = TenderInformation.LoadPackageInfo(tenderWordReader);

                for (int i = 0; i < PackageInformations.Count; i++)
                {
                    PackageInformations[i].PurchaseInformations =
                        TenderInformation.LoadPurchaseInfo(purchaseWordReader, i + 1);
                }
            }, null, (o, args) =>
            {
                Dispatcher.Invoke(() =>
                {
                    PackageList.ItemsSource = PackageInformations;
                    AutoRead.IsEnabled = true;
                    AutoRead.Content = "自动读取";
                });
            });
        }

        private void CurrentFile_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            currentPage = 1;
            UpdateContent();
        }

        #endregion

        #region Tools

        private void StartBackgroundWork(DoWorkEventHandler doWork, ProgressChangedEventHandler progressChanged, RunWorkerCompletedEventHandler completed)
        {
            var work = new BackgroundWorker { WorkerReportsProgress = true };
            work.DoWork += doWork;
            work.ProgressChanged += progressChanged;
            work.RunWorkerCompleted += completed;
            work.RunWorkerAsync();
        }

        private void LoadRunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                ShowInfoMessage("加载完成");
                StartRead.IsEnabled = true;
                StartRead.Content = "读取";

            });
        }

        private void CreateRunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                ShowInfoMessage("创建完成", () => OpenPath(ExcelPath.Text));
                CreateExcel.IsEnabled = true;
                CreateExcel.Content = "创建";
            });
        }

        private void OpenPath(string path)
        {
            Process p = new Process
            {
                StartInfo =
                {
                    FileName = "explorer.exe",
                    Arguments = $@" /select, {path}"
                }
            };
            p.Start();
        }

        private void ShowInfoMessage(string msg, Action onOkAction = null)
        {
            if (MessageBox.Show(msg, "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
            {
                onOkAction?.Invoke();
            }
        }

        private void ShowErrorMessage(string msg)
        {
            MessageBox.Show(msg, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        #endregion

        #region Word&Excel

        private void LoadWordInfo(object sender, DoWorkEventArgs e)
        {
            if (!string.IsNullOrEmpty(tenderDocText))
            {
                try
                {
                    //释放之前资源
                    tenderWordReader?.Close();

                    tenderWordReader = new WordReader(tenderDocText);
                    tenderWordReader.StartRead();
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
                    purchaseWordReader?.Close();

                    purchaseWordReader = new WordReader(purchaseDocText);
                    purchaseWordReader.StartRead();
                }
                catch (Exception exception)
                {
                    Dispatcher.Invoke(() => { ShowErrorMessage(exception.Message); });
                }
            }

            Dispatcher.Invoke(UpdateContent);
        }

        private void UpdateContent()
        {
            if (CurrentFile.SelectedValue.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") == "招标文件")
            {
                if (currentPage != 0 && tenderWordReader != null)
                    WordContent.Text = tenderWordReader.ReadPage(currentPage);
            }
            if (CurrentFile.SelectedValue.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") == "购买名单")
            {
                if (currentPage != 0 && purchaseWordReader != null)
                    WordContent.Text = purchaseWordReader.ReadPage(currentPage);
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

        private int GetContentPageCount()
        {
            if (CurrentFile.SelectedValue.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") == "招标文件")
            {
                if (tenderWordReader != null)
                    return tenderWordReader.GetPageCount();
            }
            if (CurrentFile.SelectedValue.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") == "购买名单")
            {
                if (purchaseWordReader != null)
                    return purchaseWordReader.GetPageCount();
            }

            return 0;
        }

        #endregion

        private void Window_Closed(object sender, System.EventArgs e)
        {
            tenderWordReader?.Close();
            purchaseWordReader?.Close();

            tenderWordReader = null;
            purchaseWordReader = null;
        }
    }
}
