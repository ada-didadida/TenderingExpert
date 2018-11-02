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

            InformationGrid.DataContext = Information;
        }

        public TenderingInformation Information { get; set; } = new TenderingInformation();
        public WordReader WordReader;

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
                WordReader = new WordReader(wordPath);
                WordReader.StartRead();

                Information.LoadInfo(WordReader);
            }
        }

        private void Window_Closed(object sender, System.EventArgs e)
        {
            WordReader.Close();
        }
    }
}
