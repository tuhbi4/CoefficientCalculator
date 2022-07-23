using CoefficientCalculator.Entities;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CoefficientCalculator.Services
{
    public class FileService
    {
        public FileInfo OutputFileInfo { get; }

        public List<MatchRecord> MatchRecords { get; }

        public List<CoefficientRecord> FirstCoefficientRecords { get; }

        public List<CoefficientRecord> SecondCoefficientRecords { get; }

        private readonly string baseFilePath;
        private readonly string coefficientFilePath;
        private FileInfo baseFile;
        private FileInfo coefficientFile;
        private bool isTemporaryBaseFileCreated;
        private bool isTemporaryCoefficientFileCreated;

        public FileService(string baseFilePath, string coefficientFilePath, string outputFileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.baseFilePath = baseFilePath;
            this.coefficientFilePath = coefficientFilePath;
            CreateTemporaryFilesIfNeeded();
            this.OutputFileInfo = new FileInfo(CopyXLSX(coefficientFile, outputFileName));
            MatchRecords = GetBaseCollection();
            FirstCoefficientRecords = GetCoefficientCollection(0);
            SecondCoefficientRecords = GetCoefficientCollection(1);
        }

        public void SearchForFirstCoefficientCollection()
        {
            using (var p = new ExcelPackage(OutputFileInfo))
            {
                var p1ws1 = p.Workbook.Worksheets[0];
                p1ws1.Cells[1, 4].Value = "Всего";
                p1ws1.Cells[1, 5].Value = "Ниже";
                p1ws1.Cells[1, 6].Value = "Выше";
                p1ws1.Cells[1, 7].Value = $"{baseFile.Name}";
                p1ws1.Cells[1, 8].Value = "Ниже";
                p1ws1.Cells[1, 9].Value = "Выше";
                int row = 1;

                foreach (var coefficientRecord in FirstCoefficientRecords)
                {
                    SearchResult lowerCoefficients = new SearchResult();
                    SearchResult higherCoefficients = new SearchResult();
                    SearchResult totalCoefficients = new SearchResult();
                    row++;
                    List<MatchRecord> matchRecords = MatchRecords.FindAll(x => x.G == coefficientRecord.A && x.H == coefficientRecord.B);

                    decimal.TryParse(new Regex(@" - ").Split(new Regex(@"[.]").Replace(coefficientRecord.C, ",")).ToList()[0], out decimal coefficient);

                    foreach (var matchRecord in matchRecords)
                    {
                        int matchState = GetMatchState(matchRecord.Score);

                        if (matchState == -2)
                        {
                            continue;
                        }

                        if (matchRecord.K <= coefficient)
                        {
                            if (matchState == 1)
                            {
                                lowerCoefficients.Wins++;
                                lowerCoefficients.Coefficient += matchRecord.K;
                            }
                            else
                            {
                                lowerCoefficients.Losses++;
                            }

                            lowerCoefficients.Total++;
                        }
                        else
                        {
                            if (matchState == 1)
                            {
                                higherCoefficients.Wins++;
                                higherCoefficients.Coefficient += matchRecord.K;
                            }
                            else
                            {
                                higherCoefficients.Losses++;
                            }

                            higherCoefficients.Total++;
                        }

                        totalCoefficients.Wins = lowerCoefficients.Wins + higherCoefficients.Wins;
                        totalCoefficients.Losses = lowerCoefficients.Losses + higherCoefficients.Losses;
                        totalCoefficients.Total = lowerCoefficients.Total + higherCoefficients.Total;
                        totalCoefficients.Coefficient = lowerCoefficients.Coefficient + higherCoefficients.Coefficient;

                        if (totalCoefficients.Total > 0)
                        {
                            p1ws1.Cells[row, 7].Value = $"+{totalCoefficients.Wins}-{totalCoefficients.Losses}={totalCoefficients.Total} кф {totalCoefficients.Coefficient}";
                        }

                        if (lowerCoefficients.Total > 0)
                        {
                            p1ws1.Cells[row, 8].Value = $"+{lowerCoefficients.Wins}-{lowerCoefficients.Losses}={lowerCoefficients.Total} кф {lowerCoefficients.Coefficient}";
                        }

                        if (higherCoefficients.Total > 0)
                        {
                            p1ws1.Cells[row, 9].Value = $"+{higherCoefficients.Wins}-{higherCoefficients.Losses}={higherCoefficients.Total} кф {higherCoefficients.Coefficient}";
                        }
                    }
                }

                p.Save();
            }
        }

        private void CreateTemporaryFilesIfNeeded()
        {
            baseFile = new FileInfo(baseFilePath);
            coefficientFile = new FileInfo(coefficientFilePath);

            if (baseFile.Extension == ".xls")
            {
                baseFile = new FileInfo(ConvertXLS_XLSX(new FileInfo(baseFilePath)));
                isTemporaryBaseFileCreated = true;
            }

            if (coefficientFile.Extension == ".xls")
            {
                coefficientFile = new FileInfo(ConvertXLS_XLSX(new FileInfo(coefficientFilePath)));
                isTemporaryCoefficientFileCreated = true;
            }
        }

        /// <summary>
        /// Using Microsoft.Office.Interop to convert XLS to XLSX format, to work with EPPlus library
        /// </summary>
        /// <param name="file"></param>
        private string ConvertXLS_XLSX(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application
            {
                DisplayAlerts = false
            };
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();

            return xlsxFile;
        }

        private string CopyXLSX(FileInfo file, string newName)
        {
            var newfilePath = file.DirectoryName + $"\\{newName}.xlsx";
            File.Copy(file.FullName, newfilePath, true);

            return newfilePath;
        }

        private void DeleteTemporaryFiles()
        {
            if (isTemporaryBaseFileCreated)
            {
                DeleteFile(baseFile);
            }

            if (isTemporaryCoefficientFileCreated)
            {
                DeleteFile(coefficientFile);
            }
        }

        private void DeleteFile(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private List<MatchRecord> GetBaseCollection()
        {
            List<MatchRecord> matchRecords = new List<MatchRecord>();

            using (var package = new ExcelPackage(baseFile))
            {
                ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[0];

                for (int row = 1; row < excelWorksheet.Rows.EndRow; row++)
                {
                    if (excelWorksheet.Cells[row, 7].Value != null
                        && excelWorksheet.Cells[row, 8].Value != null
                        && excelWorksheet.Cells[row, 9].Value != null
                        && excelWorksheet.Cells[row, 10].Value != null
                        && excelWorksheet.Cells[row, 11].Value != null)
                    {
                        MatchRecord matchRecord = new MatchRecord();
                        int.TryParse(excelWorksheet.Cells[row, 1].Value.ToString(), out int id);
                        int.TryParse(excelWorksheet.Cells[row, 2].Value.ToString(), out int group);
                        matchRecord.Id = id;
                        matchRecord.Group = group;
                        matchRecord.Date = excelWorksheet.Cells[row, 3].Value.ToString();
                        matchRecord.TeamOne = excelWorksheet.Cells[row, 4].Value.ToString();
                        matchRecord.TeamTwo = excelWorksheet.Cells[row, 5].Value.ToString();
                        matchRecord.Score = excelWorksheet.Cells[row, 6].Value.ToString();
                        int.TryParse(excelWorksheet.Cells[row, 7].Value.ToString(), out int fieldG);
                        matchRecord.G = fieldG;
                        int.TryParse(excelWorksheet.Cells[row, 8].Value.ToString(), out int fieldH);
                        matchRecord.H = fieldH;
                        int.TryParse(excelWorksheet.Cells[row, 9].Value.ToString(), out int fieldI);
                        matchRecord.I = fieldI;
                        int.TryParse(excelWorksheet.Cells[row, 10].Value.ToString(), out int fieldJ);
                        matchRecord.J = fieldJ;
                        decimal.TryParse(new Regex(@"[.]").Replace(excelWorksheet.Cells[row, 11].Value.ToString(), ","), out decimal fieldK);
                        matchRecord.K = fieldK;
                        decimal.TryParse(new Regex(@"[.]").Replace(excelWorksheet.Cells[row, 12].Value.ToString(), ","), out decimal fieldL);
                        matchRecord.L = fieldL;
                        decimal.TryParse(new Regex(@"[.]").Replace(excelWorksheet.Cells[row, 13].Value.ToString(), ","), out decimal fieldM);
                        matchRecord.M = fieldM;
                        matchRecords.Add(matchRecord);
                    }
                }
            }

            return matchRecords;
        }

        private List<CoefficientRecord> GetCoefficientCollection(int sheetIndex)
        {
            List<CoefficientRecord> coefficientCollection = new List<CoefficientRecord>();

            using (var package = new ExcelPackage(coefficientFile))
            {
                ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[sheetIndex];

                for (int row = 2; row < excelWorksheet.Rows.EndRow; row++)
                {
                    if (excelWorksheet.Cells[row, 1].Value != null
                        && excelWorksheet.Cells[row, 2].Value != null
                        && excelWorksheet.Cells[row, 3].Value != null)
                    {
                        CoefficientRecord coefficientRecord = new CoefficientRecord();
                        int.TryParse(excelWorksheet.Cells[row, 1].Value.ToString(), out int fieldA);
                        coefficientRecord.A = fieldA;
                        int.TryParse(excelWorksheet.Cells[row, 2].Value.ToString(), out int fieldB);
                        coefficientRecord.B = fieldB;
                        coefficientRecord.C = excelWorksheet.Cells[row, 3].Value.ToString();
                        coefficientCollection.Add(coefficientRecord);
                    }
                }
            }

            return coefficientCollection;
        }

        private int GetMatchState(string score)
        {
            Regex regex = new Regex(":");
            var stringList = regex.Split(score).ToList();
            List<int> list = new List<int>();
            stringList.ForEach(item =>
            {
                int.TryParse(item, out int number);
                list.Add(number);
            });
            int result = 0;

            switch (result)
            {
                case 0 when list.Count != 2:
                    return -2;

                case 0 when list[0] < list[1]:
                    return -1;

                case 0 when list[0] == list[1]:
                    return 0;

                case 0 when list[0] > list[1]:
                    return 1;

                default:
                    return -2;
            }
        }
    }
}