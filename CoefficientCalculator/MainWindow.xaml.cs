using CoefficientCalculator.Entities;
using CoefficientCalculator.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private readonly BackgroundWorker backgroundWorker = new BackgroundWorker();
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

            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.ProgressChanged += ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        private void BtnOpenBaseFile_Click(object sender, RoutedEventArgs e)
        {
            if (openBaseFileDialog.ShowDialog() == true)
            {
                tbBaseFile.Text = openBaseFileDialog.FileName;
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
                PrepareForWorker();
                backgroundWorker.RunWorkerAsync(1);
            }
        }

        private void BtnX_Click(object sender, RoutedEventArgs e)
        {
            if (IsValidFilePaths())
            {
                PrepareForWorker();
                backgroundWorker.RunWorkerAsync(2);
            }
        }

        private void BtnP2_Click(object sender, RoutedEventArgs e)
        {
            if (IsValidFilePaths())
            {
                PrepareForWorker();
                backgroundWorker.RunWorkerAsync(3);
            }
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

        private void CreateAndWriteData(object sender, DoWorkEventArgs e)
        {
            string outputfileName;
            int coefficientPosition = (int)e.Argument;

            switch (coefficientPosition)
            {
                case 1:
                    outputfileName = "Результат П1";
                    break;

                case 2:
                    outputfileName = "Результат Х";
                    break;

                case 3:
                    outputfileName = "Результат П2";
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(e), $"{nameof(e.Argument)} can be only 1, 2 or 3");
            }

            backgroundWorker.ReportProgress(0);
            int startColumnNumber = StartColumnNumber;
            int currentColumnNumber = StartColumnNumber + 3;
            FileInfo outputFile = FileService.CopyXLSX(new FileInfo(coefficientFilePath), outputfileName);
            backgroundWorker.ReportProgress(10);
            List<CoefficientRecord> firstCoefficientRecords = FileService.GetCoefficientRecords(new FileInfo(coefficientFilePath), 0);
            backgroundWorker.ReportProgress(30);
            List<CoefficientRecord> secondCoefficientRecords = FileService.GetCoefficientRecords(new FileInfo(coefficientFilePath), 1);
            backgroundWorker.ReportProgress(50);
            var totalFirstWorksheetSearchResultRecords = new List<SearchResultRecord>();
            var totalSecondWorksheetSearchResultRecords = new List<SearchResultRecord>();
            int progressPart = 50 / baseFilePaths.Count;
            int currentProgress = 50;

            foreach (string filePath in baseFilePaths)
            {
                currentProgress += progressPart;
                FileInfo baseFile = new FileInfo(filePath);
                bool temporaryFileCreated = FileService.CreateTemporaryFileIfNeeded(baseFile, out FileInfo newBaseFile);

                if (temporaryFileCreated)
                {
                    baseFile = newBaseFile;
                }

                List<MatchRecord> matchRecords = FileService.GetMatchRecords(baseFile);

                var firstWorksheetSearchResultRecords = FileService.GetSearchResultRecords(matchRecords, firstCoefficientRecords, coefficientPosition, 1);
                FileService.WriteSearchResults(firstWorksheetSearchResultRecords, new FileInfo(filePath), outputFile, 0, currentColumnNumber);

                var secondWorksheetSearchResultRecords = FileService.GetSearchResultRecords(matchRecords, secondCoefficientRecords, coefficientPosition, 2);
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
                backgroundWorker.ReportProgress(currentProgress);
            }

            FileService.WriteSearchResults(totalFirstWorksheetSearchResultRecords, outputFile, 0, startColumnNumber);
            FileService.WriteSearchResults(totalSecondWorksheetSearchResultRecords, outputFile, 1, startColumnNumber);
        }

        private void PrepareForWorker()
        {
            progressBar.Visibility = Visibility.Visible;
            btnPanel.IsEnabled = false;
            backgroundWorker.DoWork += CreateAndWriteData;
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Value = progressBar.Maximum;
            MessageBox.Show("Готово!");
            progressBar.Value = 0;
            progressBar.Visibility = Visibility.Hidden;
            btnPanel.IsEnabled = true;
            backgroundWorker.DoWork -= CreateAndWriteData;
        }
    }
}