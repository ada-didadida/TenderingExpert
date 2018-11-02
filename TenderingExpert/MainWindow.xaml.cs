using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
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

            InformationGrid.DataContext = Information;
        }

        public TenderingInformation Information { get; set; } = new TenderingInformation();

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
                reader.StartRead();

                Information.LoadInfo(reader);
            }
        }
    }
}
