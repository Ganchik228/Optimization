using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Optimizations
{
    public class ExcelService
    {
        public static async Task<List<Discipline>> ReadExcelDataAsync(string filePath, string? worksheetName = null)
        {
            var disciplines = new List<Discipline>();
            var parsingStats = new Dictionary<string, int>
            {
                ["Общее количество строк"] = 0,
                ["Успешно обработано"] = 0,
                ["Пропущено: пустое название"] = 0,
                ["Пропущено: пустые значения"] = 0,
                ["Пропущено: неверный формат чисел"] = 0,
                ["Пропущено: неверный семестр"] = 0,
                ["Пропущено: прочие ошибки"] = 0
            };

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Файл не найден: {filePath}");
            }

            ExcelPackage.License.SetNonCommercialPersonal("<My Name>");
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("В Excel файле нет рабочих листов");
                }

                if (string.IsNullOrEmpty(worksheetName) && package.Workbook.Worksheets.Count > 1)
                {
                    throw new InvalidOperationException("MultipleSheets");
                }

                var worksheet = string.IsNullOrEmpty(worksheetName) 
                    ? package.Workbook.Worksheets[0] 
                    : package.Workbook.Worksheets[worksheetName];

                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Лист '{worksheetName}' не найден");
                }

                if (worksheet.Dimension == null)
                {
                    throw new InvalidOperationException("Рабочий лист пуст");
                }

                int totalRows = worksheet.Dimension.Rows;
                int totalColumns = worksheet.Dimension.Columns;
                parsingStats["Общее количество строк"] = totalRows - 1;

                if (totalRows < 2)
                {
                    throw new InvalidOperationException("В файле недостаточно данных (нужно минимум 2 строки)");
                }

                if (totalColumns < 5)
                {
                    throw new InvalidOperationException("В файле недостаточно столбцов (нужно минимум 5 столбцов)");
                }

                var columnHeaders = new[]
                {
                    worksheet.Cells[1, 1].Text,
                    worksheet.Cells[1, 2].Text,
                    worksheet.Cells[1, 3].Text,
                    worksheet.Cells[1, 4].Text,
                    worksheet.Cells[1, 5].Text
                };

                var columnExamples = new List<string>[5];
                for (int i = 0; i < 5; i++)
                {
                    columnExamples[i] = new List<string>();
                }

                for (int rowIndex = 2; rowIndex <= Math.Min(totalRows, 6); rowIndex++)
                {
                    for (int colIndex = 1; colIndex <= 5; colIndex++)
                    {
                        var cellValue = worksheet.Cells[rowIndex, colIndex]?.Text ?? "";
                        columnExamples[colIndex - 1].Add(cellValue);
                    }
                }

                Console.WriteLine("Заголовки столбцов:");
                for (int columnIndex = 0; columnIndex < columnHeaders.Length; columnIndex++)
                {
                    Console.WriteLine($"Столбец {columnIndex + 1}: {columnHeaders[columnIndex]}");
                }

                var tasks = new List<Task>();
                var disciplineIndex = 0;
                var indexLock = new object();
                for (int rowIndex = 2; rowIndex <= totalRows; rowIndex++)
                {
                    var currentRowIndex = rowIndex;
                    tasks.Add(Task.Run(() =>
                    {
                        try
                        {
                            if (currentRowIndex > totalRows)
                            {
                                Console.WriteLine($"Пропуск строки {currentRowIndex}: строка за пределами диапазона");
                                return;
                            }

                            var disciplineName = worksheet.Cells[currentRowIndex, 1]?.Text;
                            if (string.IsNullOrWhiteSpace(disciplineName))
                            {
                                Console.WriteLine($"Пропуск строки {currentRowIndex}: пустое название дисциплины");
                                lock (parsingStats) { parsingStats["Пропущено: пустое название"]++; }
                                return;
                            }

                            var minWorkloadText = worksheet.Cells[currentRowIndex, 2]?.Text;
                            var maxWorkloadText = worksheet.Cells[currentRowIndex, 3]?.Text;
                            var significanceText = worksheet.Cells[currentRowIndex, 4]?.Text;
                            var semesterText = worksheet.Cells[currentRowIndex, 5]?.Text;

                            if (string.IsNullOrWhiteSpace(minWorkloadText) ||
                                string.IsNullOrWhiteSpace(maxWorkloadText) ||
                                string.IsNullOrWhiteSpace(significanceText) ||
                                string.IsNullOrWhiteSpace(semesterText))
                            {
                                Console.WriteLine($"Пропуск строки {currentRowIndex}: пустые значения в ячейках");
                                lock (parsingStats) { parsingStats["Пропущено: пустые значения"]++; }
                                return;
                            }

                            double minWorkload, maxWorkload, significanceCoefficient;
                            int semesterNumber;

                            try
                            {
                                minWorkload = double.Parse(minWorkloadText.Replace(",", "."), CultureInfo.InvariantCulture);
                                maxWorkload = double.Parse(maxWorkloadText.Replace(",", "."), CultureInfo.InvariantCulture);
                                significanceCoefficient = double.Parse(significanceText.Replace(",", "."), CultureInfo.InvariantCulture);
                                semesterNumber = (int)double.Parse(semesterText.Replace(",", "."), CultureInfo.InvariantCulture);
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine($"Пропуск строки {currentRowIndex}: некорректный формат числовых значений");
                                lock (parsingStats) { parsingStats["Пропущено: неверный формат чисел"]++; }
                                return;
                            }

                            if (semesterNumber < 1 || semesterNumber > 8)
                            {
                                Console.WriteLine($"Пропуск строки {currentRowIndex}: некорректный номер семестра ({semesterNumber})");
                                lock (parsingStats) { parsingStats["Пропущено: неверный семестр"]++; }
                                return;
                            }

                            int currentDisciplineIndex;
                            lock (indexLock)
                            {
                                currentDisciplineIndex = disciplineIndex++;
                            }

                            var newDiscipline = new Discipline
                            {
                                Name = disciplineName,
                                MinWorkload = minWorkload,
                                MaxWorkload = maxWorkload,
                                SignificanceCoefficient = significanceCoefficient,
                                Semester = semesterNumber,
                                Index = currentDisciplineIndex
                            };

                            lock (disciplines)
                            {
                                disciplines.Add(newDiscipline);
                                parsingStats["Успешно обработано"]++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ошибка при обработке строки {currentRowIndex}: {ex.Message}");
                            lock (parsingStats) { parsingStats["Пропущено: прочие ошибки"]++; }
                        }
                    }));
                }
                await Task.WhenAll(tasks);
            }

            if (disciplines.Count == 0)
            {
                throw new InvalidOperationException("Не удалось прочитать ни одной дисциплины из файла");
            }

            Console.WriteLine($"\nУспешно прочитано дисциплин: {disciplines.Count}");
            
            ShowParsingInfo(disciplines, parsingStats);
            
            return disciplines;
        }

        public static async Task SaveResultsToExcelAsync(string filePath, List<Discipline> disciplines, Dictionary<string, double> bestVariant, List<(Dictionary<string, double> variant, double objective)> topVariants)
        {
            ExcelPackage.License.SetNonCommercialPersonal("<My Name>");
            using (var excelPackage = new ExcelPackage())
            {
                var resultsWorksheet = excelPackage.Workbook.Worksheets.Add("Результаты");
                var sortedDisciplines = disciplines.OrderBy(discipline => discipline.Semester).ToList();
                int currentRow = 1;
                int variantsToSave = Math.Min(topVariants.Count, 50);

                currentRow = WriteVariantTable(resultsWorksheet, sortedDisciplines, bestVariant,
                    "ЛУЧШИЙ ВАРИАНТ", currentRow, Color.Yellow);

                currentRow += 2;
                for (int variantIndex = 1; variantIndex < variantsToSave; variantIndex++)
                {
                    var variant = topVariants[variantIndex];
                    string variantTitle = $"ВАРИАНТ {variantIndex}";
                    Color highlightColor = Color.LightGray;

                    currentRow = WriteVariantTable(resultsWorksheet, sortedDisciplines, variant.variant,
                        variantTitle, currentRow, highlightColor);

                    currentRow += 2;
                }

                resultsWorksheet.Cells[resultsWorksheet.Dimension.Address].AutoFitColumns();
                await excelPackage.SaveAsAsync(new FileInfo(filePath));
            }
        }

        private static int WriteVariantTable(ExcelWorksheet worksheet, List<Discipline> sortedDisciplines, 
            Dictionary<string, double> variant, string title, int startRow, Color highlightColor)
        {
            int currentRow = startRow;

            worksheet.Cells[currentRow, 1].Value = title;
            worksheet.Cells[currentRow, 1, currentRow, 3].Merge = true;
            using (var titleRange = worksheet.Cells[currentRow, 1, currentRow, 3])
            {
                titleRange.Style.Font.Bold = true;
                titleRange.Style.Font.Size = 12;
                titleRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                titleRange.Style.Fill.BackgroundColor.SetColor(highlightColor);
                titleRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "Название дисциплины";
            worksheet.Cells[currentRow, 2].Value = "Семестр";
            worksheet.Cells[currentRow, 3].Value = "Трудоемкость";
            
            using (var headerRange = worksheet.Cells[currentRow, 1, currentRow, 3])
            {
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            currentRow++;

            var semesterSums = new Dictionary<int, double>();
            for (int sem = 1; sem <= 8; sem++)
            {
                semesterSums[sem] = sortedDisciplines
                    .Where(d => d.Semester == sem)
                    .Sum(d => variant[d.UniqueName]);
            }

            var yearSums = new Dictionary<int, double>();
            for (int year = 1; year <= 4; year++)
            {
                int sem1 = (year - 1) * 2 + 1;
                int sem2 = (year - 1) * 2 + 2;
                yearSums[year] = semesterSums[sem1] + semesterSums[sem2];
            }

            int currentSemester = -1;
            int currentYear = -1;

            foreach (var discipline in sortedDisciplines)
            {
                int disciplineYear = (discipline.Semester + 1) / 2;
                if (disciplineYear != currentYear)
                {
                    currentYear = disciplineYear;
                    
                    worksheet.Cells[currentRow, 1].Value = $"ГОД {currentYear}";
                    worksheet.Cells[currentRow, 2].Value = "";
                    worksheet.Cells[currentRow, 3].Value = yearSums[currentYear];
                    
                    using (var yearRange = worksheet.Cells[currentRow, 1, currentRow, 3])
                    {
                        yearRange.Style.Font.Bold = true;
                        yearRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        yearRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(173, 216, 230));
                    }
                    currentRow++;
                }

                if (discipline.Semester != currentSemester)
                {
                    currentSemester = discipline.Semester;
                    
                    worksheet.Cells[currentRow, 1].Value = $"  Семестр {currentSemester}";
                    worksheet.Cells[currentRow, 2].Value = "";
                    worksheet.Cells[currentRow, 3].Value = semesterSums[currentSemester];
                    
                    using (var semesterRange = worksheet.Cells[currentRow, 1, currentRow, 3])
                    {
                        semesterRange.Style.Font.Bold = true;
                        semesterRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        semesterRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(230, 230, 250));
                    }
                    currentRow++;
                }

                worksheet.Cells[currentRow, 1].Value = $"    {discipline.Name}";
                worksheet.Cells[currentRow, 2].Value = discipline.Semester;
                worksheet.Cells[currentRow, 3].Value = variant[discipline.UniqueName];
                currentRow++;
            }

            var totalSum = yearSums.Values.Sum();
            worksheet.Cells[currentRow, 1].Value = "ИТОГО:";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = totalSum;
            
            using (var totalRange = worksheet.Cells[currentRow, 1, currentRow, 3])
            {
                totalRange.Style.Font.Bold = true;
                totalRange.Style.Font.Size = 12;
                totalRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                totalRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 215, 0));
            }
            currentRow++;

            return currentRow;
        }

        public static void ShowParsingInfo(List<Discipline> disciplines, Dictionary<string, int> parsingStats)
        {
            var semesterGroups = disciplines.GroupBy(d => d.Semester).OrderBy(g => g.Key);
            var fixedDisciplines = disciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) < 0.001).Count();
            var variableDisciplines = disciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) > 0.001).Count();
            
            var infoText = new StringBuilder();
            infoText.AppendLine("ИНФОРМАЦИЯ О ЗАГРУЖЕННЫХ ДАННЫХ:");
            infoText.AppendLine();
            
            infoText.AppendLine("СТАТИСТИКА ПАРСИНГА:");
            foreach (var stat in parsingStats)
            {
                infoText.AppendLine($"{stat.Key}: {stat.Value}");
            }
            infoText.AppendLine();
            
            infoText.AppendLine($"Всего дисциплин: {disciplines.Count}");
            infoText.AppendLine($"Фиксированных дисциплин: {fixedDisciplines}");
            infoText.AppendLine($"Переменных дисциплин: {variableDisciplines}");
            infoText.AppendLine();
            
            infoText.AppendLine("РАСПРЕДЕЛЕНИЕ ПО СЕМЕСТРАМ:");
            foreach (var group in semesterGroups)
            {
                var semesterFixed = group.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) < 0.001).Count();
                var semesterVariable = group.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) > 0.001).Count();
                infoText.AppendLine($"Семестр {group.Key}: {group.Count()} дисциплин (фикс: {semesterFixed}, вар: {semesterVariable})");
            }
            infoText.AppendLine();
            
            infoText.AppendLine("ДИАПАЗОНЫ ТРУДОЕМКОСТИ:");
            var minWorkload = disciplines.Min(d => d.MinWorkload);
            var maxWorkload = disciplines.Max(d => d.MaxWorkload);
            var avgSignificance = disciplines.Average(d => d.SignificanceCoefficient);
            
            infoText.AppendLine($"Минимальная трудоемкость: {minWorkload:F1}");
            infoText.AppendLine($"Максимальная трудоемкость: {maxWorkload:F1}");
            infoText.AppendLine($"Средний коэффициент значимости: {avgSignificance:F3}");
            
            MessageBox.Show(infoText.ToString(), "Информация о загруженных данных", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
