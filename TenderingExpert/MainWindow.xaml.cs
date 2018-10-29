using System;
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
        }

        private void SelectFile_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                Filter = "Word File(*.doc, *.docx) |*doc;*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
                WordPath.Text = openFileDialog.FileName;
        }

        private void StartRead_OnClick(object sender, RoutedEventArgs e)
        {
            var wordPath = WordPath.Text;
            if (!string.IsNullOrEmpty(wordPath))
            {
                WordReader reader = new WordReader(wordPath);
                try
                {
                    reader.StartRead();

                    TenderingInformation information = new TenderingInformation(reader);
                }
                catch (Exception exception)
                {
                    Result.Text = exception.Message;
                }
            }
        }
    }
}
