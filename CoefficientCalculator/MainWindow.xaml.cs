using CoefficientCalculator.Entities;
using CoefficientCalculator.Services;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace CoefficientCalculator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int StartColumnNumber = 4;
        private readonly OpenFileDialog openBaseFileDialog;
        private readonly OpenFileDialog openCoefficientFileDialog;
        private string baseFilePath;
        private List<string> baseFilePaths;
        private string coefficientFilePath;

        public MainWindow()
        {
            openBaseFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel files (*.xlsx, *.xls)|*.xlsx; *.xls"
            };

            baseFilePaths = new List<string>();

            openCoefficientFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx, *.xls)|*.xlsx; *.xls"
            };

            InitializeComponent();
        }

        private void BtnOpenBaseFile_Click(object sender, RoutedEventArgs e)
        {
            if (openBaseFileDialog.ShowDialog() == true)
            {
                tbBaseFile.Text = openBaseFileDialog.FileName;
                baseFilePath = openBaseFileDialog.FileName;
                baseFilePaths = openBaseFileDialog.FileNames.ToList();
            }
        }

        private void BtnOpenCoefficientFile_Click(object sender, RoutedEventArgs e)
        {
            if (openCoefficientFileDialog.ShowDialog() == true)
            {
                tbCoefficientFile.Text = openCoefficientFileDialog.FileName;
                coefficientFilePath = openCoefficientFileDialog.FileName;
            }
        }

        private void BtnP1_Click(object sender, RoutedEventArgs e)
        {
            if (IsValidFilePaths())
            {
                int startColumnNumber = StartColumnNumber;
                int currentColumnNumber = StartColumnNumber + 3;
                FileInfo outputFile = FileService.CopyXLSX(new FileInfo(coefficientFilePath), "П1");
                List<CoefficientRecord> firstCoefficientRecords = FileService.GetCoefficientRecords(new FileInfo(coefficientFilePath), 0);
                List<CoefficientRecord> secondCoefficientRecords = FileService.GetCoefficientRecords(new FileInfo(coefficientFilePath), 1);
                var totalFirstWorksheetSearchResultRecords = new List<SearchResultRecord>();
                var totalSecondWorksheetSearchResultRecords = new List<SearchResultRecord>();

                foreach (string filePath in baseFilePaths)
                {
                    FileInfo baseFile = new FileInfo(filePath);
                    bool temporaryFileCreated = FileService.CreateTemporaryFileIfNeeded(baseFile, out FileInfo newBaseFile);

                    if (temporaryFileCreated)
                    {
                        baseFile = newBaseFile;
                    }

                    List<MatchRecord> matchRecords = FileService.GetMatchRecords(baseFile);

                    var firstWorksheetSearchResultRecords = FileService.GetSearchResultRecords(matchRecords, firstCoefficientRecords, 1, 1);
                    FileService.WriteSearchResults(firstWorksheetSearchResultRecords, new FileInfo(filePath), outputFile, 0, currentColumnNumber);

                    var secondWorksheetSearchResultRecords = FileService.GetSearchResultRecords(matchRecords, secondCoefficientRecords, 1, 2);
                    FileService.WriteSearchResults(secondWorksheetSearchResultRecords, new FileInfo(filePath), outputFile, 1, currentColumnNumber);

                    if (currentColumnNumber == StartColumnNumber + 3)
                    {
                        totalFirstWorksheetSearchResultRecords.AddRange(firstWorksheetSearchResultRecords);
                        totalSecondWorksheetSearchResultRecords.AddRange(secondWorksheetSearchResultRecords);
                    }
                    else
                    {
                        FileService.CalculateTotalResult(totalFirstWorksheetSearchResultRecords, firstWorksheetSearchResultRecords);
                        FileService.CalculateTotalResult(totalSecondWorksheetSearchResultRecords, secondWorksheetSearchResultRecords);
                    }

                    if (temporaryFileCreated)
                    {
                        FileService.DeleteFile(newBaseFile);
                    }

                    currentColumnNumber += 3;
                }

                FileService.WriteSearchResults(totalFirstWorksheetSearchResultRecords, outputFile, 0, startColumnNumber);
                FileService.WriteSearchResults(totalSecondWorksheetSearchResultRecords, outputFile, 1, startColumnNumber);

                MessageBox.Show("Готово!");
            }
        }

        private void BtnX_Click(object sender, RoutedEventArgs e)
        {
            IsValidFilePaths();
        }

        private void BtnP2_Click(object sender, RoutedEventArgs e)
        {
            IsValidFilePaths();
        }

        private bool IsValidFilePaths()
        {
            if (string.IsNullOrEmpty(openBaseFileDialog.FileName))
            {
                MessageBox.Show("Сначала выберите исходный файл.");
                return false;
            }
            else if (string.IsNullOrEmpty(openCoefficientFileDialog.FileName))
            {
                MessageBox.Show("Сначала выберите файл коэффициентов.");
                return false;
            }

            return true;
        }
    }
}