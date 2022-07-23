using CoefficientCalculator.Entities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CoefficientCalculator.Services
{
    public static class FileService
    {
        public static bool CreateTemporaryFileIfNeeded(FileInfo file, out FileInfo newFile)
        {
            if (file.Extension == ".xls")
            {
                newFile = new FileInfo(ConvertXLS_XLSX(file));
                return true;
            }

            newFile = null;
            return false;
        }

        public static FileInfo CopyXLSX(FileInfo file, string newName)
        {
            var newfilePath = file.DirectoryName + $"\\{newName}.xlsx";
            File.Copy(file.FullName, newfilePath, true);

            return new FileInfo(newfilePath);
        }

        public static void DeleteFile(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        public static List<MatchRecord> GetMatchRecords(FileInfo baseFile)
        {
            List<MatchRecord> matchRecords = new List<MatchRecord>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(baseFile))
            {
                ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[0];

                for (int row = 1; row < excelWorksheet.Rows.EndRow; row++)
                {
                    if (excelWorksheet.Cells[row, 11].Value != null
                        && excelWorksheet.Cells[row, 12].Value != null
                        && excelWorksheet.Cells[row, 13].Value != null)
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

        public static List<CoefficientRecord> GetCoefficientRecords(FileInfo coefficientFile, int sheetIndex)
        {
            List<CoefficientRecord> coefficientCollection = new List<CoefficientRecord>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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

        public static List<SearchResultRecord> GetSearchResultRecords(List<MatchRecord> matchRecords,
            List<CoefficientRecord> coefficientRecords, int coefficientPosition, int coefficientSheetNumber)
        {
            var searchResultRecords = new List<SearchResultRecord>();
            List<MatchRecord> selectedMatchRecords;

            foreach (var coefficientRecord in coefficientRecords)
            {
                switch (coefficientSheetNumber)
                {
                    case 1:
                        selectedMatchRecords = matchRecords.FindAll(x => x.G == coefficientRecord.A && x.H == coefficientRecord.B);
                        break;

                    case 2:
                        selectedMatchRecords = matchRecords.FindAll(x => x.I == coefficientRecord.A && x.J == coefficientRecord.B);
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(nameof(coefficientSheetNumber), $"{nameof(coefficientSheetNumber)} can be only 1 or 2");
                }

                searchResultRecords.Add(GetSearchResultRecord(coefficientRecord, selectedMatchRecords, coefficientPosition));
            }

            return searchResultRecords;
        }

        public static void CalculateTotalResult(List<SearchResultRecord> totalSearchResultRecords,
            List<SearchResultRecord> searchResultRecords)
        {
            for (int i = 0; i < searchResultRecords.Count; i++)
            {
                totalSearchResultRecords[i].LowerCoefficients.Wins += searchResultRecords[i].LowerCoefficients.Wins;
                totalSearchResultRecords[i].LowerCoefficients.Losses += searchResultRecords[i].LowerCoefficients.Losses;
                totalSearchResultRecords[i].LowerCoefficients.Total += searchResultRecords[i].LowerCoefficients.Total;
                totalSearchResultRecords[i].LowerCoefficients.Coefficient += searchResultRecords[i].LowerCoefficients.Coefficient;
                totalSearchResultRecords[i].HigherCoefficients.Wins += searchResultRecords[i].HigherCoefficients.Wins;
                totalSearchResultRecords[i].HigherCoefficients.Losses += searchResultRecords[i].HigherCoefficients.Losses;
                totalSearchResultRecords[i].HigherCoefficients.Total += searchResultRecords[i].HigherCoefficients.Total;
                totalSearchResultRecords[i].HigherCoefficients.Coefficient += searchResultRecords[i].HigherCoefficients.Coefficient;
                totalSearchResultRecords[i].TotalCoefficients.Wins += searchResultRecords[i].TotalCoefficients.Wins;
                totalSearchResultRecords[i].TotalCoefficients.Losses += searchResultRecords[i].TotalCoefficients.Losses;
                totalSearchResultRecords[i].TotalCoefficients.Total += searchResultRecords[i].TotalCoefficients.Total;
                totalSearchResultRecords[i].TotalCoefficients.Coefficient += searchResultRecords[i].TotalCoefficients.Coefficient;
            }
        }

        public static void WriteSearchResults(List<SearchResultRecord> searchResultRecords,
            FileInfo baseFileInfo, FileInfo outputFileInfo, int worksheetIndex, int startColumnNumber)
        {
            WriteSearchResults(searchResultRecords, baseFileInfo.Name, outputFileInfo, worksheetIndex, startColumnNumber);
        }

        public static void WriteSearchResults(List<SearchResultRecord> searchResultRecords,
            FileInfo outputFileInfo, int worksheetIndex, int startColumnNumber)
        {
            WriteSearchResults(searchResultRecords, string.Empty, outputFileInfo, worksheetIndex, startColumnNumber);
        }

        private static void WriteSearchResults(List<SearchResultRecord> searchResultRecords,
                            string fileName, FileInfo outputFileInfo, int worksheetIndex, int startColumnNumber)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var p = new ExcelPackage(outputFileInfo))
            {
                var p1ws1 = p.Workbook.Worksheets[worksheetIndex];

                if (string.IsNullOrEmpty(fileName))
                {
                    p1ws1.Cells[1, startColumnNumber].Value = $"Всего";
                }
                else
                {
                    p1ws1.Cells[1, startColumnNumber].Value = $"{fileName}";
                }

                p1ws1.Cells[1, startColumnNumber + 1].Value = "Ниже коэф";
                p1ws1.Cells[1, startColumnNumber + 2].Value = "Выше коэф";

                var startRowNumber = 2;
                foreach (var searchResultRecord in searchResultRecords)
                {
                    if (searchResultRecord.TotalCoefficients.Total > 0)
                    {
                        p1ws1.Cells[startRowNumber, startColumnNumber].Value =
                            $"+{searchResultRecord.TotalCoefficients.Wins}" +
                            $"-{searchResultRecord.TotalCoefficients.Losses}" +
                            $"={searchResultRecord.TotalCoefficients.Total}" +
                            $" кф {searchResultRecord.TotalCoefficients.Coefficient}";
                    }

                    if (searchResultRecord.LowerCoefficients.Total > 0)
                    {
                        p1ws1.Cells[startRowNumber, startColumnNumber + 1].Value =
                            $"+{searchResultRecord.LowerCoefficients.Wins}" +
                            $"-{searchResultRecord.LowerCoefficients.Losses}" +
                            $"={searchResultRecord.LowerCoefficients.Total}" +
                            $" кф {searchResultRecord.LowerCoefficients.Coefficient}";
                    }

                    if (searchResultRecord.HigherCoefficients.Total > 0)
                    {
                        p1ws1.Cells[startRowNumber, startColumnNumber + 2].Value =
                            $"+{searchResultRecord.HigherCoefficients.Wins}" +
                            $"-{searchResultRecord.HigherCoefficients.Losses}" +
                            $"={searchResultRecord.HigherCoefficients.Total}" +
                            $" кф {searchResultRecord.HigherCoefficients.Coefficient}";
                    }

                    startRowNumber++;
                }

                p.Save();
            }
        }

        private static SearchResultRecord GetSearchResultRecord(CoefficientRecord coefficientRecord,
                                        List<MatchRecord> selectedMatchRecords, int coefficientPosition)
        {
            var lowerCoefficients = new SearchResult();
            var higherCoefficients = new SearchResult();
            var totalCoefficients = new SearchResult();
            int targetMatchState;

            switch (coefficientPosition)
            {
                case 1:
                    targetMatchState = 1;
                    break;

                case 2:
                    targetMatchState = 0;
                    break;

                case 3:
                    targetMatchState = -1;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(coefficientPosition), $"{nameof(coefficientPosition)} can be only 1, 2 or 3");
            }

            decimal.TryParse(new Regex(@" - ").Split(new Regex(@"[.]").Replace(coefficientRecord.C, ","))
                .ToList()[coefficientPosition - 1], out decimal coefficient);

            foreach (var matchRecord in selectedMatchRecords)
            {
                int matchState = GetMatchState(matchRecord.Score);

                if (matchState == -2)
                {
                    continue;
                }

                var matchCoefficient = GetCoefficientFromRecord(matchRecord, coefficientPosition);

                if (matchCoefficient <= coefficient)
                {
                    if (matchState == targetMatchState)
                    {
                        lowerCoefficients.Wins++;
                        lowerCoefficients.Coefficient += matchCoefficient;
                    }
                    else
                    {
                        lowerCoefficients.Losses++;
                    }

                    lowerCoefficients.Total++;
                }
                else
                {
                    if (matchState == targetMatchState)
                    {
                        higherCoefficients.Wins++;
                        higherCoefficients.Coefficient += matchCoefficient;
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
            }

            var searchResultRecord = new SearchResultRecord
            {
                LowerCoefficients = lowerCoefficients,
                HigherCoefficients = higherCoefficients,
                TotalCoefficients = totalCoefficients
            };

            return searchResultRecord;
        }

        /// <summary>
        /// Using Microsoft.Office.Interop to convert XLS to XLSX format, to work with EPPlus library
        /// </summary>
        /// <param name="file"></param>
        private static string ConvertXLS_XLSX(FileInfo file)
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

        private static int GetMatchState(string score)
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
                    throw new ArgumentException($"Unhandled error with {nameof(list)}");
            }
        }

        private static decimal GetCoefficientFromRecord(MatchRecord matchRecord, int coefficientPosition)
        {
            switch (coefficientPosition)
            {
                case 1:
                    return matchRecord.K;

                case 2:
                    return matchRecord.L;

                case 3:
                    return matchRecord.M;

                default:
                    throw new ArgumentOutOfRangeException(nameof(coefficientPosition), $"{nameof(coefficientPosition)} can be only 1, 2 or 3");
            }
        }
    }
}