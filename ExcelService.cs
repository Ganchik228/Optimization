using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Optimizations
{
    public class ExcelService
    {
        public static async Task<List<Discipline>> ReadExcelDataAsync(string filePath, string? worksheetName = null)
        {
            var disciplines = new List<Discipline>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Файл не найден: {filePath}");
            }

            ExcelPackage.License.SetNonCommercialPersonal("<My Name>"); // Ensure this is your actual license name or handle appropriately
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("В Excel файле нет рабочих листов");
                }

                ExcelWorksheet? worksheet;
                if (string.IsNullOrEmpty(worksheetName))
                {
                    if (package.Workbook.Worksheets.Count > 1)
                    {
                        // Let Form1 handle sheet selection
                        throw new InvalidOperationException("MultipleSheets");
                    }
                    worksheet = package.Workbook.Worksheets[0];
                }
                else
                {
                    worksheet = package.Workbook.Worksheets[worksheetName];
                }

                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Лист '{worksheetName ?? "Первый лист"}' не найден");
                }

                if (worksheet.Dimension == null)
                {
                    throw new InvalidOperationException("Рабочий лист пуст");
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                if (rowCount < 2)
                {
                    throw new InvalidOperationException("В файле недостаточно данных (нужно минимум 2 строки)");
                }

                if (colCount < 5)
                {
                    throw new InvalidOperationException("В файле недостаточно столбцов (нужно минимум 5 столбцов)");
                }
                
                // Optional: Log headers if needed for debugging
                // var headers = new[]
                // {
                //     worksheet.Cells[1, 1].Text,
                //     worksheet.Cells[1, 2].Text,
                //     worksheet.Cells[1, 3].Text,
                //     worksheet.Cells[1, 4].Text,
                //     worksheet.Cells[1, 5].Text
                // };
                // Console.WriteLine("Заголовки столбцов: " + string.Join(", ", headers));

                var tasks = new List<Task>();
                for (int row = 2; row <= rowCount; row++)
                {
                    var currentRow = row; // Capture row variable for closure
                    tasks.Add(Task.Run(() =>
                    {
                        try
                        {
                            // Bounds check, though worksheet.Cells should handle out of range gracefully by returning null
                            if (currentRow > worksheet.Dimension.Rows) 
                            {
                                Console.WriteLine($"Пропуск строки {currentRow}: строка за пределами диапазона");
                                return;
                            }

                            var name = worksheet.Cells[currentRow, 1]?.Text;
                            if (string.IsNullOrWhiteSpace(name))
                            {
                                Console.WriteLine($"Пропуск строки {currentRow}: пустое название дисциплины");
                                return;
                            }

                            var minWorkloadStr = worksheet.Cells[currentRow, 2]?.Text;
                            var maxWorkloadStr = worksheet.Cells[currentRow, 3]?.Text;
                            var significanceStr = worksheet.Cells[currentRow, 4]?.Text;
                            var semesterStr = worksheet.Cells[currentRow, 5]?.Text;

                            if (string.IsNullOrWhiteSpace(minWorkloadStr) ||
                                string.IsNullOrWhiteSpace(maxWorkloadStr) ||
                                string.IsNullOrWhiteSpace(significanceStr) ||
                                string.IsNullOrWhiteSpace(semesterStr))
                            {
                                Console.WriteLine($"Пропуск строки {currentRow} ({name}): пустые значения в ячейках");
                                return;
                            }

                            double minWorkload, maxWorkload, significance;
                            int semester;

                            try
                            {
                                minWorkload = double.Parse(minWorkloadStr.Replace(",", "."), CultureInfo.InvariantCulture);
                                maxWorkload = double.Parse(maxWorkloadStr.Replace(",", "."), CultureInfo.InvariantCulture);
                                significance = double.Parse(significanceStr.Replace(",", "."), CultureInfo.InvariantCulture);
                                semester = (int)double.Parse(semesterStr.Replace(",", "."), CultureInfo.InvariantCulture); // Or int.Parse if always integer
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine($"Пропуск строки {currentRow} ({name}): некорректный формат числовых значений");
                                return;
                            }

                            if (semester < 1 || semester > 8) // Assuming 8 semesters max
                            {
                                Console.WriteLine($"Пропуск строки {currentRow} ({name}): некорректный номер семестра ({semester})");
                                return;
                            }

                            var discipline = new Discipline(name) // Pass name to constructor
                            {
                                MinWorkload = minWorkload,
                                MaxWorkload = maxWorkload,
                                SignificanceCoefficient = significance,
                                Semester = semester
                            };

                            lock (disciplines) // Ensure thread-safe add to list
                            {
                                disciplines.Add(discipline);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log detailed error, perhaps with row number and cell content if possible
                            Console.WriteLine($"Ошибка при обработке строки {currentRow}: {ex.Message}");
                        }
                    }));
                }
                await Task.WhenAll(tasks);
            }

            if (disciplines.Count == 0)
            {
                // This might be a valid scenario if the file is empty or all rows are skipped.
                // Consider if this should be an exception or just an empty list.
                // For now, matching Avtomat's behavior:
                throw new InvalidOperationException("Не удалось прочитать ни одной дисциплины из файла. Проверьте формат данных и консоль на наличие ошибок.");
            }

            Console.WriteLine($"\nУспешно прочитано дисциплин: {disciplines.Count}");
            return disciplines;
        }

        public static async Task SaveInitialDataAsSingleVariantAsync(string filePath, List<Discipline> initialDisciplines)
        {
            ExcelPackage.License.SetNonCommercialPersonal("<My Name>"); // Ensure this is your actual license name or handle appropriately
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Исходные данные (1 вариант)");
                int currentRow = 1;

                var sortedDisciplines = initialDisciplines.OrderBy(d => d.Semester).ThenBy(d => d.Name).ToList();

                // Header for the "single variant"
                worksheet.Cells[currentRow, 1].Value = "Исходные данные (сохранены как один вариант)";
                worksheet.Cells[currentRow, 1, currentRow, 4].Merge = true; // Span 4 columns
                worksheet.Cells[currentRow, 1].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 1].Style.Font.Size = 14;
                worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[currentRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                currentRow++;

                // Column Headers
                worksheet.Cells[currentRow, 1].Value = "Название дисциплины";
                worksheet.Cells[currentRow, 2].Value = "Семестр";
                worksheet.Cells[currentRow, 3].Value = "Мин. трудоемкость";
                worksheet.Cells[currentRow, 4].Value = "Макс. трудоемкость";
                
                using (var range = worksheet.Cells[currentRow, 1, currentRow, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                }
                currentRow++;

                // Data for the initial disciplines
                foreach (var discipline in sortedDisciplines)
                {
                    worksheet.Cells[currentRow, 1].Value = discipline.Name;
                    worksheet.Cells[currentRow, 2].Value = discipline.Semester;
                    worksheet.Cells[currentRow, 3].Value = discipline.MinWorkload;
                    worksheet.Cells[currentRow, 4].Value = discipline.MaxWorkload;
                    currentRow++;
                }
                
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                await package.SaveAsAsync(new FileInfo(filePath));
            }
        }

        public static async Task SaveResultsToExcelAsync(string filePath, List<Discipline> disciplines, Dictionary<string, double> bestVariant, List<Dictionary<string, double>> topVariants)
        {
            ExcelPackage.License.SetNonCommercialPersonal("<My Name>"); 
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Результаты");
                int currentRow = 1;

                var sortedDisciplines = disciplines.OrderBy(d => d.Semester).ThenBy(d => d.Name).ToList();

                int maxVariantsToSave = Math.Min(topVariants.Count, 100); 

                for (int i = 0; i < maxVariantsToSave; i++)
                {
                    // topVariants is now List<Dictionary<string, double>>
                    var variantDict = topVariants[i]; 

                    // Variant Header - removed objective
                    worksheet.Cells[currentRow, 1].Value = $"Вариант {i + 1}";
                    worksheet.Cells[currentRow, 1, currentRow, 3].Merge = true; 
                    worksheet.Cells[currentRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 1].Style.Font.Size = 14;
                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    
                    // Highlight the first variant (considered "best" by order of generation/selection)
                    if (i == 0 && variantDict == bestVariant) // Check if it's the designated bestVariant
                    {
                        worksheet.Cells[currentRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gold);
                    }
                    currentRow++;

                    // Column Headers for this variant block
                    worksheet.Cells[currentRow, 1].Value = "Название дисциплины";
                    worksheet.Cells[currentRow, 2].Value = "Семестр";
                    worksheet.Cells[currentRow, 3].Value = "Трудоемкость";
                    
                    using (var range = worksheet.Cells[currentRow, 1, currentRow, 3])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    }
                    currentRow++;

                    // Data for this variant
                    foreach (var discipline in sortedDisciplines)
                    {
                        worksheet.Cells[currentRow, 1].Value = discipline.Name;
                        worksheet.Cells[currentRow, 2].Value = discipline.Semester;

                        if (variantDict.TryGetValue(discipline.UniqueName, out double workload))
                        {
                            worksheet.Cells[currentRow, 3].Value = workload;
                        }
                        else
                        {
                            worksheet.Cells[currentRow, 3].Value = "N/A"; 
                        }
                        currentRow++;
                    }
                    
                    currentRow++; 
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                await package.SaveAsAsync(new FileInfo(filePath));
            }
        }
    }
}
