using CoefficientCalculator.Services;
using Microsoft.Win32;
using System.Windows;

namespace CoefficientCalculator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly OpenFileDialog openFileDialog;
        //private readonly List<string> fileNames = new List<string>();
        private FileService fileService;
        private string baseFilePath;
        private string coefficientFilePath;

        public MainWindow()
        {
            openFileDialog = new OpenFileDialog();
            InitializeComponent();
        }

        private void BtnOpenBaseFile_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx; *.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                tbBaseFile.Text = openFileDialog.FileName;
                baseFilePath = openFileDialog.FileName;
            }
        }

        private void BtnOpenCoefficientFile_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx; *.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                tbCoefficientFile.Text = openFileDialog.FileName;
                coefficientFilePath = openFileDialog.FileName;
            }
        }

        private void BtnP1_Click(object sender, RoutedEventArgs e)
        {
            if (IsValidFilePaths())
            {
                fileService = new FileService(baseFilePath, coefficientFilePath, "П1");

                fileService.SearchForFirstCoefficientCollection();
                fileService.WriteFirstCoefficientCollectionSearchResults(0, 7);
                fileService.SearchForSecondCoefficientCollection();
                fileService.WriteSecondCoefficientCollectionSearchResults(1, 7);

                MessageBox.Show("Готово!");
            }
        }

        private void BtnX_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(baseFilePath))
            {
                MessageBox.Show("Сначала выберите исходный файл.");
            }
            else
            {
                MessageBox.Show("Готово!");
            }
        }

        private void BtnP2_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(baseFilePath))
            {
                MessageBox.Show("Сначала выберите исходный файл.");
            }
            else
            {
                MessageBox.Show("Готово!");
            }
        }

        private bool IsValidFilePaths()
        {
            if (string.IsNullOrEmpty(baseFilePath))
            {
                MessageBox.Show("Сначала выберите исходный файл.");
                return false;
            }
            else if (string.IsNullOrEmpty(coefficientFilePath))
            {
                MessageBox.Show("Сначала выберите файл коэффициентов.");
                return false;
            }

            return true;
        }
    }
}