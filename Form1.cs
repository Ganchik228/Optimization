using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using OfficeOpenXml;
using System.Collections.Concurrent;

namespace Optimizations
{
    public partial class Form1 : Form
    {        public Form1()
        {
            InitializeComponent();
            
            statusLabel.Text = "Готов к работе";
            progressBar.Visible = false;
            
            button3.Enabled = false;
            
            this.Icon = SystemIcons.Application;
        }public class Discipline
        {
            public string Name { get; set; } = "";
            public double MinWorkload { get; set; }
            public double MaxWorkload { get; set; }
            public double SignificanceCoefficient { get; set; }
            public int Semester { get; set; }
            public int Index { get; set; }
            public string UniqueName => $"{Name} (семестр {Semester}) - {Index}";
        }        public static async Task<List<Discipline>> ReadExcelDataAsync(string filePath, string? worksheetName = null)
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
                }                int totalRows = worksheet.Dimension.Rows;
                int totalColumns = worksheet.Dimension.Columns;
                parsingStats["Общее количество строк"] = totalRows - 1; // Исключаем заголовок

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

                // Собираем примеры данных из первых нескольких строк для анализа типов
                var columnExamples = new List<string>[5];
                for (int i = 0; i < 5; i++)
                {
                    columnExamples[i] = new List<string>();
                }

                // Берем примеры из первых 5 строк данных (не включая заголовок)
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
                            }                            var disciplineName = worksheet.Cells[currentRowIndex, 1]?.Text;
                            if (string.IsNullOrWhiteSpace(disciplineName))
                            {
                                Console.WriteLine($"Пропуск строки {currentRowIndex}: пустое название дисциплины");
                                lock (parsingStats) { parsingStats["Пропущено: пустое название"]++; }
                                return;
                            }

                            var minWorkloadText = worksheet.Cells[currentRowIndex, 2]?.Text;
                            var maxWorkloadText = worksheet.Cells[currentRowIndex, 3]?.Text;
                            var significanceText = worksheet.Cells[currentRowIndex, 4]?.Text;
                            var semesterText = worksheet.Cells[currentRowIndex, 5]?.Text;                            if (string.IsNullOrWhiteSpace(minWorkloadText) ||
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
                            }                            catch (FormatException)
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
                            };                            lock (disciplines)
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
            }            Console.WriteLine($"\nУспешно прочитано дисциплин: {disciplines.Count}");
            
            // Показываем информацию о парсинге в MessageBox
            ShowParsingInfo(disciplines, parsingStats);
            
            return disciplines;
        }
        
        public static Task<List<Dictionary<string, double>>> GenerateValidVariantsAsync(List<Discipline> disciplines)
        {
            return Task.Run(() =>
            {
                var semesterPairs = new List<(int Sem1, int Sem2, double TargetSum)>
                {
                    (1, 2, 60.5),
                    (3, 4, 59.5),
                    (5, 6, 60),
                    (7, 8, 60)
                };

                var blockVariants = new List<List<BlockVariant>>(semesterPairs.Count);
                var blockVariantsLock = new object();                Parallel.ForEach(semesterPairs, semesterPair =>
                {
                    var (firstSemester, secondSemester, targetSum) = semesterPair;
                    var blockDisciplines = disciplines.Where(discipline => discipline.Semester == firstSemester || discipline.Semester == secondSemester).ToList();
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: дисциплин {blockDisciplines.Count}");
                    var fixedDisciplines = blockDisciplines.Where(discipline => Math.Abs(discipline.MinWorkload - discipline.MaxWorkload) < 0.001).ToList();
                    var variableDisciplines = blockDisciplines.Where(discipline => Math.Abs(discipline.MinWorkload - discipline.MaxWorkload) > 0.001).ToList();
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: фиксированных {fixedDisciplines.Count}, переменных {variableDisciplines.Count}");

                    var baseVariant = fixedDisciplines.ToDictionary(discipline => discipline.UniqueName, discipline => discipline.MinWorkload);
                    var valueOptionsList = variableDisciplines
                        .Select(discipline => GeneratePossibleValues(discipline.MinWorkload, discipline.MaxWorkload))
                        .ToList();
                    if (variableDisciplines.Count > 0)
                        Console.WriteLine($"Блок {firstSemester}+{secondSemester}: всего комбинаций {valueOptionsList.Select(optionList => optionList.Count).Aggregate(1, (accumulator, count) => accumulator * count)}");

                    var localVariants = new List<BlockVariant>();
                    int rejectedVariantsCount = 0;

                    void RecurseVariants(int disciplineIndex, Dictionary<string, double> currentVariant)
                    {
                        if (disciplineIndex == variableDisciplines.Count)
                        {
                            var excludedDisciplineNames = new HashSet<string> { "Онтологическое моделирование", "Проектирование пользовательского интерфейса" };
                            double sumFirstSemester = 0, sumSecondSemester = 0, sumPair = 0, excludedSum = 0;
                            foreach (var discipline in blockDisciplines)
                            {
                                double workloadValue = baseVariant.ContainsKey(discipline.UniqueName) ? baseVariant[discipline.UniqueName] : (currentVariant.ContainsKey(discipline.UniqueName) ? currentVariant[discipline.UniqueName] : 0);
                                if (discipline.Semester == firstSemester) sumFirstSemester += workloadValue;
                                if (discipline.Semester == secondSemester) sumSecondSemester += workloadValue;
                                if (discipline.Semester == firstSemester || discipline.Semester == secondSemester)
                                {
                                    sumPair += workloadValue;
                                    if (excludedDisciplineNames.Contains(discipline.Name))
                                        excludedSum += workloadValue;
                                }
                            }

                            double sumWithoutExcluded = sumPair - excludedSum;
                            bool isValidVariant = true;
                            string rejectionReason = "";
                            if (Math.Abs(sumWithoutExcluded - targetSum) > 0.001)
                            {
                                isValidVariant = false;
                                rejectionReason = $"Сумма (без исключённых) {sumWithoutExcluded} != {targetSum}";
                            }
                            else if (Math.Abs(sumFirstSemester - sumSecondSemester) > 6)
                            {
                                isValidVariant = false;
                                rejectionReason = $"Разница между семестрами {Math.Abs(sumFirstSemester - sumSecondSemester)} > 6";
                            }

                            if (!isValidVariant)
                            {
                                if (rejectedVariantsCount < 5)
                                {
                                    var allWorkloadValues = new Dictionary<string, double>(baseVariant);
                                    foreach (var keyValue in currentVariant) allWorkloadValues[keyValue.Key] = keyValue.Value;
                                    string workloadValuesText = string.Join(", ", allWorkloadValues.Select(keyValue => $"{keyValue.Key}:{keyValue.Value}"));
                                    rejectedVariantsCount++;
                                }
                                return;
                            }

                            var validVariant = new Dictionary<string, double>(baseVariant);
                            foreach (var keyValue in currentVariant) validVariant[keyValue.Key] = keyValue.Value;

                            double objectiveFirstSemester = 0, objectiveSecondSemester = 0;
                            foreach (var discipline in blockDisciplines)
                            {
                                double workloadValue = validVariant[discipline.UniqueName];
                                if (discipline.Semester == firstSemester) objectiveFirstSemester += workloadValue * discipline.SignificanceCoefficient;
                                if (discipline.Semester == secondSemester) objectiveSecondSemester += workloadValue * discipline.SignificanceCoefficient;
                            }

                            lock (localVariants)
                            {
                                localVariants.Add(new BlockVariant
                                {
                                    Values = validVariant,
                                    SumsBySemester = new[] { objectiveFirstSemester, objectiveSecondSemester }
                                });
                            }
                            return;
                        }
                        var currentDiscipline = variableDisciplines[disciplineIndex];
                        foreach (var workloadValue in valueOptionsList[disciplineIndex])
                        {
                            currentVariant[currentDiscipline.UniqueName] = workloadValue;
                            RecurseVariants(disciplineIndex + 1, currentVariant);
                            currentVariant.Remove(currentDiscipline.UniqueName);
                        }
                    }

                    RecurseVariants(0, new Dictionary<string, double>());
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: найдено вариантов {localVariants.Count}");

                    var topTenVariants = localVariants
                        .OrderByDescending(variant => variant.SumsBySemester[0] * variant.SumsBySemester[1])
                        .Take(10)
                        .ToList();

                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: отобрано топ-10 вариантов");
                    lock (blockVariantsLock)
                    {
                        blockVariants.Add(topTenVariants);
                    }
                });                var allResultVariants = new ConcurrentBag<Dictionary<string, double>>();
                int totalBlocks = blockVariants.Count;
                var blockSizes = blockVariants.Select(blockVariantList => blockVariantList.Count).ToArray();
                long totalCombinations = blockSizes.Aggregate(1L, (accumulator, size) => accumulator * size);
                Console.WriteLine($"Итоговое число комбинаций: {totalCombinations}");

                int maxTopVariants = 100;
                var globalTopVariants = new SortedSet<(double objectiveValue, Dictionary<string, double> variant)>(Comparer<(double, Dictionary<string, double>)>.Create((firstItem, secondItem) =>
                {
                    int comparison = firstItem.Item1.CompareTo(secondItem.Item1);
                    if (comparison == 0)
                        return firstItem.Item2.GetHashCode().CompareTo(secondItem.Item2.GetHashCode());
                    return comparison;
                }));
                object topVariantsLock = new object();

                Parallel.ForEach(
                    Partitioner.Create(0L, totalCombinations),
                    () => new SortedSet<(double objectiveValue, Dictionary<string, double> variant)>(Comparer<(double, Dictionary<string, double>)>.Create((firstItem, secondItem) =>
                    {
                        int comparison = firstItem.Item1.CompareTo(secondItem.Item1);
                        if (comparison == 0)
                            return firstItem.Item2.GetHashCode().CompareTo(secondItem.Item2.GetHashCode());
                        return comparison;
                    })),
                    (combinationRange, loopState, threadLocalTopVariants) =>
                    {
                        for (long combinationIndex = combinationRange.Item1; combinationIndex < combinationRange.Item2; combinationIndex++)
                        {
                            var blockIndices = new int[totalBlocks];
                            long tempIndex = combinationIndex;
                            for (int blockIndex = totalBlocks - 1; blockIndex >= 0; blockIndex--)
                            {
                                blockIndices[blockIndex] = (int)(tempIndex % blockSizes[blockIndex]);
                                tempIndex /= blockSizes[blockIndex];
                            }

                            var combinedVariant = new Dictionary<string, double>();
                            double objectiveValue = 1.0;
                            for (int blockIndex = 0; blockIndex < totalBlocks; blockIndex++)
                            {
                                var selectedBlockVariant = blockVariants[blockIndex][blockIndices[blockIndex]];
                                foreach (var workloadPair in selectedBlockVariant.Values)
                                    combinedVariant[workloadPair.Key] = workloadPair.Value;
                                objectiveValue *= (selectedBlockVariant.SumsBySemester[0] + selectedBlockVariant.SumsBySemester[1]);
                            }

                            if (threadLocalTopVariants.Count < maxTopVariants)
                            {
                                threadLocalTopVariants.Add((objectiveValue, combinedVariant));
                            }
                            else if (objectiveValue > threadLocalTopVariants.Min.Item1)
                            {
                                threadLocalTopVariants.Remove(threadLocalTopVariants.Min);
                                threadLocalTopVariants.Add((objectiveValue, combinedVariant));
                            }
                        }
                        return threadLocalTopVariants;
                    },
                    threadLocalTopVariants =>
                    {
                        lock (topVariantsLock)
                        {
                            foreach (var topVariant in threadLocalTopVariants)
                            {
                                if (globalTopVariants.Count < maxTopVariants)
                                {
                                    globalTopVariants.Add(topVariant);
                                }
                                else if (topVariant.Item1 > globalTopVariants.Min.Item1)
                                {
                                    globalTopVariants.Remove(globalTopVariants.Min);
                                    globalTopVariants.Add(topVariant);
                                }
                            }
                        }
                    }
                );

                var topHundredVariants = globalTopVariants.Reverse().Take(100).ToList();
                return topHundredVariants.Select(variantTuple => variantTuple.Item2).ToList();
            });
        }
        private static List<double> GeneratePossibleValues(double minValue, double maxValue)
        {
            var possibleValues = new List<double>();
            for (double currentValue = minValue; currentValue <= maxValue; currentValue += 1)
            {
                possibleValues.Add(currentValue);
            }
            //Console.WriteLine($"[DEBUG] Возможные значения для диапазона {minValue}-{maxValue}: {string.Join(", ", possibleValues)}");
            return possibleValues;
        }
        public static List<Discipline>[] PrepareSemDisc(List<Discipline> disciplines)
        {
            var semesterDisciplines = new List<Discipline>[8];
            for (int semesterNumber = 1; semesterNumber <= 8; semesterNumber++)
                semesterDisciplines[semesterNumber - 1] = disciplines.Where(discipline => discipline.Semester == semesterNumber).ToList();
            return semesterDisciplines;
        }
        public static double CalculateObjectiveFast(List<Discipline>[] semesterDisciplines, Dictionary<string, double> workloadVariant)
        {
            double objectiveProduct = 1.0;
            for (int semesterIndex = 0; semesterIndex < 8; semesterIndex++)
            {
                double semesterSum = 0;
                foreach (var discipline in semesterDisciplines[semesterIndex])
                    semesterSum += workloadVariant[discipline.UniqueName] * discipline.SignificanceCoefficient;
                objectiveProduct *= semesterSum;
            }
            return objectiveProduct;
        }
          class BlockVariant
        {
            public Dictionary<string, double> Values = new Dictionary<string, double>();
            public double[] SumsBySemester = new double[2];
        }
        private async Task SaveResultsToExcelAsync(string filePath, List<Discipline> disciplines, Dictionary<string, double> bestVariant, List<(Dictionary<string, double> variant, double objective)> topVariants)
        {
            ExcelPackage.License.SetNonCommercialPersonal("<My Name>");
            using (var excelPackage = new ExcelPackage())
            {
                var resultsWorksheet = excelPackage.Workbook.Worksheets.Add("Результаты");
                var sortedDisciplines = disciplines.OrderBy(discipline => discipline.Semester).ToList();
                int currentRow = 1;
                int variantsToSave = Math.Min(topVariants.Count, 50); 
                currentRow = WriteVariantTable(resultsWorksheet, sortedDisciplines, bestVariant,
                    "ЛУЧШИЙ ВАРИАНТ", currentRow, Color.Yellow);                currentRow += 2;
                // Сохраняем остальные варианты (начиная со второго, так как первый уже сохранен как "ЛУЧШИЙ ВАРИАНТ")
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

        private int WriteVariantTable(ExcelWorksheet worksheet, List<Discipline> sortedDisciplines, 
            Dictionary<string, double> variant, string title, int startRow, Color highlightColor)
        {
            int currentRow = startRow;

            // Заголовок варианта
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

            // Заголовки столбцов
            worksheet.Cells[currentRow, 1].Value = "Название дисциплины";
            worksheet.Cells[currentRow, 2].Value = "Семестр";
            worksheet.Cells[currentRow, 3].Value = "Трудоемкость";
            
            using (var headerRange = worksheet.Cells[currentRow, 1, currentRow, 3])
            {
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            currentRow++;            // Данные по дисциплинам
            foreach (var discipline in sortedDisciplines)
            {
                worksheet.Cells[currentRow, 1].Value = discipline.Name;
                worksheet.Cells[currentRow, 2].Value = discipline.Semester;
                worksheet.Cells[currentRow, 3].Value = variant[discipline.UniqueName];
                currentRow++;
            }

            return currentRow;
        }private List<Discipline>? currentDisciplines;
        private List<(Dictionary<string, double> variant, double objective)>? currentTopVariants;
        private Dictionary<string, double>? currentBestVariant;        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
                statusLabel.Text = "Выбор файла...";
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                    openFileDialog.Title = "Выберите файл Excel с данными";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Показываем прогресс-бар и обновляем статус
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        statusLabel.Text = "Загрузка данных из Excel файла...";
                        button1.Enabled = false;
                        button3.Enabled = false;                        try
                        {
                            currentDisciplines = await ReadExcelDataAsync(openFileDialog.FileName);
                        }
                        catch (InvalidOperationException ex) when (ex.Message == "MultipleSheets")
                        {
                            statusLabel.Text = "Выбор листа Excel...";
                            using (var excelPackage = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                            {
                                var worksheetNames = excelPackage.Workbook.Worksheets.Select(worksheet => worksheet.Name).ToArray();
                                using (var sheetSelectionForm = new Form())
                                {
                                    sheetSelectionForm.Text = "Выберите лист";
                                    sheetSelectionForm.Width = 300;
                                    sheetSelectionForm.Height = 150;
                                    sheetSelectionForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                                    sheetSelectionForm.StartPosition = FormStartPosition.CenterParent;
                                    sheetSelectionForm.MaximizeBox = false;
                                    sheetSelectionForm.MinimizeBox = false;
                                    sheetSelectionForm.BackColor = Color.White;
                                    sheetSelectionForm.Font = new Font("Segoe UI", 9F);

                                    var instructionLabel = new Label
                                    {
                                        Text = "Выберите лист для загрузки:",
                                        Location = new Point(10, 10),
                                        AutoSize = true,
                                        Font = new Font("Segoe UI", 10F)
                                    };

                                    var sheetComboBox = new ComboBox
                                    {
                                        Location = new Point(10, 40),
                                        Width = 260,
                                        DropDownStyle = ComboBoxStyle.DropDownList,
                                        Font = new Font("Segoe UI", 10F)
                                    };
                                    sheetComboBox.Items.AddRange(worksheetNames);
                                    sheetComboBox.SelectedIndex = 0;

                                    var okButton = new Button
                                    {
                                        Text = "OK",
                                        DialogResult = DialogResult.OK,
                                        Location = new Point(100, 70),
                                        BackColor = Color.FromArgb(0, 123, 255),
                                        ForeColor = Color.White,
                                        FlatStyle = FlatStyle.Flat,
                                        Font = new Font("Segoe UI", 10F)
                                    };
                                    okButton.FlatAppearance.BorderSize = 0;                                    sheetSelectionForm.Controls.AddRange(new Control[] { instructionLabel, sheetComboBox, okButton });
                                    sheetSelectionForm.AcceptButton = okButton;

                                    if (sheetSelectionForm.ShowDialog() == DialogResult.OK)
                                    {
                                        var selectedSheetName = sheetComboBox.SelectedItem?.ToString();
                                        if (selectedSheetName != null)
                                        {
                                            statusLabel.Text = $"Загрузка данных из листа '{selectedSheetName}'...";
                                            currentDisciplines = await ReadExcelDataAsync(openFileDialog.FileName, selectedSheetName);
                                        }
                                    }
                                    else
                                    {
                                        progressBar.Visible = false;
                                        statusLabel.Text = "Готов к работе";
                                        button1.Enabled = true;
                                        return;
                                    }
                                }
                            }
                        }

                        if (currentDisciplines == null)
                        {
                            MessageBox.Show("Не удалось загрузить данные из файла.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            progressBar.Visible = false;
                            statusLabel.Text = "Ошибка загрузки данных";
                            button1.Enabled = true;
                            return;
                        }

                        statusLabel.Text = "Генерация вариантов оптимизации...";
                        var semesterDisciplinesArray = PrepareSemDisc(currentDisciplines);
                        var validVariants = await GenerateValidVariantsAsync(currentDisciplines);                        statusLabel.Text = "Расчет целевой функции...";
                        var scoredVariants = new ConcurrentBag<(Dictionary<string, double> variant, double objective)>();
                        Parallel.ForEach(validVariants, workloadVariant =>
                        {
                            double objectiveValue = CalculateObjectiveFast(semesterDisciplinesArray, workloadVariant);
                            scoredVariants.Add((workloadVariant, objectiveValue));
                        });

                        statusLabel.Text = "Сортировка результатов...";
                        currentTopVariants = scoredVariants
                            .OrderByDescending(variantTuple => variantTuple.objective)
                            .Take(10000)
                            .ToList();

                        if (currentTopVariants.Count > 0)
                        {
                            var bestVariantTuple = currentTopVariants.First();
                            currentBestVariant = bestVariantTuple.Item1;                            Console.WriteLine("\nЛучший допустимый вариант:");
                            Console.WriteLine(string.Join(", ", bestVariantTuple.Item1.Select(keyValue => $"{keyValue.Key}:{keyValue.Value}")));
                            
                            statusLabel.Text = "Заполнение таблицы результатов...";
                            dataGridView1.Rows.Clear();

                            var sortedDisciplines = currentDisciplines.OrderBy(discipline => discipline.Semester).ToList();
  
                            foreach (var discipline in sortedDisciplines)
                            {
                                var disciplineWorkload = bestVariantTuple.Item1[discipline.UniqueName];
                                dataGridView1.Rows.Add(
                                    discipline.Name,
                                    disciplineWorkload.ToString("F1"),
                                    discipline.SignificanceCoefficient.ToString("F2"),
                                    discipline.Semester
                                );
                            }
                            
                            statusLabel.Text = $"Готово! Найдено {currentTopVariants.Count} вариантов оптимизации";
                            button3.Enabled = true;
                        }
                        else
                        {
                            MessageBox.Show("Допустимых вариантов не найдено.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            statusLabel.Text = "Допустимых вариантов не найдено";
                        }
                        
                        // Скрываем прогресс-бар и активируем кнопки
                        progressBar.Visible = false;
                        button1.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                progressBar.Visible = false;
                statusLabel.Text = "Ошибка выполнения операции";
                button1.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }        private async void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (currentDisciplines == null || currentTopVariants == null || currentBestVariant == null)
                {
                    MessageBox.Show("Сначала загрузите и обработайте данные!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (SaveFileDialog saveResultsDialog = new SaveFileDialog())
                {
                    saveResultsDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                    saveResultsDialog.Title = "Сохранить результаты";
                    saveResultsDialog.DefaultExt = "xlsx";
                    saveResultsDialog.AddExtension = true;

                    if (saveResultsDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Показываем прогресс
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        statusLabel.Text = "Сохранение результатов в Excel файл...";
                        button3.Enabled = false;
                        button1.Enabled = false;
                          int variantsToSave = Math.Min(currentTopVariants.Count, 50);
                        await SaveResultsToExcelAsync(saveResultsDialog.FileName, currentDisciplines, currentBestVariant, currentTopVariants);
                        
                        // Скрываем прогресс
                        progressBar.Visible = false;
                        statusLabel.Text = $"Результаты сохранены! Файл: {Path.GetFileName(saveResultsDialog.FileName)}";
                        button3.Enabled = true;
                        button1.Enabled = true;
                        
                        MessageBox.Show($"Результаты успешно сохранены!\nСохранено вариантов: {variantsToSave}\nФайл: {saveResultsDialog.FileName}", 
                            "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                statusLabel.Text = "Ошибка сохранения файла";
                button3.Enabled = true;
                button1.Enabled = true;
                MessageBox.Show($"Произошла ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }private static void ShowParsingInfo(List<Discipline> disciplines, Dictionary<string, int> parsingStats)
        {
            var semesterGroups = disciplines.GroupBy(d => d.Semester).OrderBy(g => g.Key);
            var fixedDisciplines = disciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) < 0.001).Count();
            var variableDisciplines = disciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) > 0.001).Count();
            
            var infoText = new StringBuilder();
            infoText.AppendLine("ИНФОРМАЦИЯ О ЗАГРУЖЕННЫХ ДАННЫХ:");
            infoText.AppendLine();
            
            // Статистика парсинга
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
