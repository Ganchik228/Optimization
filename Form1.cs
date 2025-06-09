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
    {
        public Form1()
        {
            InitializeComponent();
        }        public class Discipline
        {
            public string Name { get; set; } = "";
            public double MinWorkload { get; set; }
            public double MaxWorkload { get; set; }
            public double SignificanceCoefficient { get; set; }
            public int Semester { get; set; }
            public int Index { get; set; }
            public string UniqueName => $"{Name} (семестр {Semester}) - {Index}";
        }

        public static async Task<List<Discipline>> ReadExcelDataAsync(string filePath, string? worksheetName = null)
        {
            var disciplines = new List<Discipline>();

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

                var headers = new[]
                {
                    worksheet.Cells[1, 1].Text,
                    worksheet.Cells[1, 2].Text,
                    worksheet.Cells[1, 3].Text,
                    worksheet.Cells[1, 4].Text,
                    worksheet.Cells[1, 5].Text
                };

                Console.WriteLine("Заголовки столбцов:");
                for (int i = 0; i < headers.Length; i++)
                {
                    Console.WriteLine($"Столбец {i + 1}: {headers[i]}");
                }                var tasks = new List<Task>();
                var disciplineIndex = 0;
                var indexLock = new object();
                
                for (int row = 2; row <= rowCount; row++)
                {
                    var currentRow = row;
                    tasks.Add(Task.Run(() =>
                    {
                        try
                        {
                            if (currentRow > rowCount)
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
                                Console.WriteLine($"Пропуск строки {currentRow}: пустые значения в ячейках");
                                return;
                            }

                            double minWorkload, maxWorkload, significance;
                            int semester;

                            try
                            {
                                minWorkload = double.Parse(minWorkloadStr.Replace(",", "."), CultureInfo.InvariantCulture);
                                maxWorkload = double.Parse(maxWorkloadStr.Replace(",", "."), CultureInfo.InvariantCulture);
                                significance = double.Parse(significanceStr.Replace(",", "."), CultureInfo.InvariantCulture);
                                semester = (int)double.Parse(semesterStr.Replace(",", "."), CultureInfo.InvariantCulture);
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine($"Пропуск строки {currentRow}: некорректный формат числовых значений");
                                return;
                            }

                            if (semester < 1 || semester > 8)
                            {
                                Console.WriteLine($"Пропуск строки {currentRow}: некорректный номер семестра ({semester})");
                                return;
                            }

                            int currentIndex;
                            lock (indexLock)
                            {
                                currentIndex = disciplineIndex++;
                            }

                            var discipline = new Discipline
                            {
                                Name = name,
                                MinWorkload = minWorkload,
                                MaxWorkload = maxWorkload,
                                SignificanceCoefficient = significance,
                                Semester = semester,
                                Index = currentIndex
                            };

                            lock (disciplines)
                            {
                                disciplines.Add(discipline);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ошибка при обработке строки {currentRow}: {ex.Message}");
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
                var blockVariantsLock = new object();

                Parallel.ForEach(semesterPairs, pair =>
                {
                    var (sem1, sem2, targetSum) = pair;
                    var blockDisciplines = disciplines.Where(d => d.Semester == sem1 || d.Semester == sem2).ToList();
                    Console.WriteLine($"Блок {sem1}+{sem2}: дисциплин {blockDisciplines.Count}");
                    var fixedBlock = blockDisciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) < 0.001).ToList();
                    var variableBlock = blockDisciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) > 0.001).ToList();
                    Console.WriteLine($"Блок {sem1}+{sem2}: фиксированных {fixedBlock.Count}, переменных {variableBlock.Count}");

                    var baseVariant = fixedBlock.ToDictionary(d => d.UniqueName, d => d.MinWorkload);
                    var valueOptions = variableBlock
                        .Select(d => GeneratePossibleValues(d.MinWorkload, d.MaxWorkload))
                        .ToList();
                    if (variableBlock.Count > 0)
                        Console.WriteLine($"Блок {sem1}+{sem2}: всего комбинаций {valueOptions.Select(l => l.Count).Aggregate(1, (a, b) => a * b)}");

                    var localVariants = new List<BlockVariant>();
                    int rejectedCount = 0;

                    void Recurse(int idx, Dictionary<string, double> current)
                    {
                        if (idx == variableBlock.Count)
                        {
                            var excludedNames = new HashSet<string> { "Онтологическое моделирование", "Проектирование пользовательского интерфейса" };
                            double sumSem1 = 0, sumSem2 = 0, sumPair = 0, excludedSum = 0;
                            foreach (var d in blockDisciplines)
                            {
                                double v = baseVariant.ContainsKey(d.UniqueName) ? baseVariant[d.UniqueName] : (current.ContainsKey(d.UniqueName) ? current[d.UniqueName] : 0);
                                if (d.Semester == sem1) sumSem1 += v;
                                if (d.Semester == sem2) sumSem2 += v;
                                if (d.Semester == sem1 || d.Semester == sem2)
                                {
                                    sumPair += v;
                                    if (excludedNames.Contains(d.Name))
                                        excludedSum += v;
                                }
                            }

                            double sumWithoutExcluded = sumPair - excludedSum;
                            bool ok = true;
                            string reason = "";
                            if (Math.Abs(sumWithoutExcluded - targetSum) > 0.001)
                            {
                                ok = false;
                                reason = $"Сумма (без исключённых) {sumWithoutExcluded} != {targetSum}";
                            }
                            else if (Math.Abs(sumSem1 - sumSem2) > 6)
                            {
                                ok = false;
                                reason = $"Разница между семестрами {Math.Abs(sumSem1 - sumSem2)} > 6";
                            }

                            if (!ok)
                            {
                                if (rejectedCount < 5)
                                {
                                    var allVals = new Dictionary<string, double>(baseVariant);
                                    foreach (var kv in current) allVals[kv.Key] = kv.Value;
                                    string values = string.Join(", ", allVals.Select(kv => $"{kv.Key}:{kv.Value}"));
                                    rejectedCount++;
                                }
                                return;
                            }

                            var variant = new Dictionary<string, double>(baseVariant);
                            foreach (var kv in current) variant[kv.Key] = kv.Value;

                            double objSem1 = 0, objSem2 = 0;
                            foreach (var d in blockDisciplines)
                            {
                                double v = variant[d.UniqueName];
                                if (d.Semester == sem1) objSem1 += v * d.SignificanceCoefficient;
                                if (d.Semester == sem2) objSem2 += v * d.SignificanceCoefficient;
                            }

                            lock (localVariants)
                            {
                                localVariants.Add(new BlockVariant
                                {
                                    Values = variant,
                                    SumsBySemester = new[] { objSem1, objSem2 }
                                });
                            }
                            return;
                        }
                        var disc = variableBlock[idx];
                        foreach (var value in valueOptions[idx])
                        {
                            current[disc.UniqueName] = value;
                            Recurse(idx + 1, current);
                            current.Remove(disc.UniqueName);
                        }
                    }

                    Recurse(0, new Dictionary<string, double>());
                    Console.WriteLine($"Блок {sem1}+{sem2}: найдено вариантов {localVariants.Count}");

                    var top10 = localVariants
                        .OrderByDescending(v => v.SumsBySemester[0] * v.SumsBySemester[1])
                        .Take(10)
                        .ToList();

                    Console.WriteLine($"Блок {sem1}+{sem2}: отобрано топ-10 вариантов");
                    lock (blockVariantsLock)
                    {
                        blockVariants.Add(top10);
                    }
                });

                var allResults = new ConcurrentBag<Dictionary<string, double>>();
                int blockCount = blockVariants.Count;
                var blockLengths = blockVariants.Select(b => b.Count).ToArray();
                long totalComb = blockLengths.Aggregate(1L, (a, b) => a * b);
                Console.WriteLine($"Итоговое число комбинаций: {totalComb}");

                int maxTop = 100;
                var globalTop = new SortedSet<(double, Dictionary<string, double>)>(Comparer<(double, Dictionary<string, double>)>.Create((a, b) =>
                {
                    int cmp = a.Item1.CompareTo(b.Item1);
                    if (cmp == 0)
                        return a.Item2.GetHashCode().CompareTo(b.Item2.GetHashCode());
                    return cmp;
                }));
                object topLock = new object();

                Parallel.ForEach(
                    Partitioner.Create(0L, totalComb),
                    () => new SortedSet<(double, Dictionary<string, double>)>(Comparer<(double, Dictionary<string, double>)>.Create((a, b) =>
                    {
                        int cmp = a.Item1.CompareTo(b.Item1);
                        if (cmp == 0)
                            return a.Item2.GetHashCode().CompareTo(b.Item2.GetHashCode());
                        return cmp;
                    })),
                    (range, state, localTop) =>
                    {
                        for (long idx = range.Item1; idx < range.Item2; idx++)
                        {
                            var indices = new int[blockCount];
                            long t = idx;
                            for (int i = blockCount - 1; i >= 0; i--)
                            {
                                indices[i] = (int)(t % blockLengths[i]);
                                t /= blockLengths[i];
                            }

                            var variant = new Dictionary<string, double>();
                            double obj = 1.0;
                            for (int b = 0; b < blockCount; b++)
                            {
                                var blockVar = blockVariants[b][indices[b]];
                                foreach (var kv in blockVar.Values)
                                    variant[kv.Key] = kv.Value;
                                obj *= (blockVar.SumsBySemester[0] + blockVar.SumsBySemester[1]);
                            }

                            if (localTop.Count < maxTop)
                            {
                                localTop.Add((obj, variant));
                            }
                            else if (obj > localTop.Min.Item1)
                            {
                                localTop.Remove(localTop.Min);
                                localTop.Add((obj, variant));
                            }
                        }
                        return localTop;
                    },
                    localTop =>
                    {
                        lock (topLock)
                        {
                            foreach (var item in localTop)
                            {
                                if (globalTop.Count < maxTop)
                                {
                                    globalTop.Add(item);
                                }
                                else if (item.Item1 > globalTop.Min.Item1)
                                {
                                    globalTop.Remove(globalTop.Min);
                                    globalTop.Add(item);
                                }
                            }
                        }
                    }
                );

                var top100 = globalTop.Reverse().Take(100).ToList();
                return top100.Select(x => x.Item2).ToList();
            });
        }
        
        private static List<double> GeneratePossibleValues(double min, double max)
        {
            var values = new List<double>();
            for (double v = min; v <= max; v += 1)
            {
                values.Add(v);
            }
            //Console.WriteLine($"[DEBUG] Возможные значения для диапазона {min}-{max}: {string.Join(", ", values)}");
            return values;
        }
        
        public static List<Discipline>[] PrepareSemDisc(List<Discipline> disciplines)
        {
            var semDisc = new List<Discipline>[8];
            for (int sem = 1; sem <= 8; sem++)
                semDisc[sem - 1] = disciplines.Where(d => d.Semester == sem).ToList();
            return semDisc;
        }

        public static double CalculateObjectiveFast(List<Discipline>[] semDisc, Dictionary<string, double> variant)
        {
            double product = 1.0;
            for (int sem = 0; sem < 8; sem++)
            {
                double sum = 0;
                foreach (var disc in semDisc[sem])
                    sum += variant[disc.UniqueName] * disc.SignificanceCoefficient;
                product *= sum;
            }
            return product;
        }
          class BlockVariant
        {
            public Dictionary<string, double> Values = new Dictionary<string, double>();
            public double[] SumsBySemester = new double[2];
        }

        private async Task SaveResultsToExcelAsync(string filePath, List<Discipline> disciplines, Dictionary<string, double> bestVariant, List<(Dictionary<string, double> variant, double objective)> topVariants)
        {
            ExcelPackage.License.SetNonCommercialPersonal("<My Name>");
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Результаты");

                worksheet.Cells[1, 1].Value = "Название дисциплины";
                worksheet.Cells[1, 2].Value = "Семестр";
                worksheet.Cells[1, 3].Value = "Трудоемкость (лучший вариант)";

                for (int i = 0; i < Math.Min(topVariants.Count, 10000); i++)
                {
                    worksheet.Cells[1, i + 4].Value = $"Вариант {i + 1}";
                }

                using (var range = worksheet.Cells[1, 1, 1, Math.Min(topVariants.Count, 10000) + 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }

                var sortedDisciplines = disciplines.OrderBy(d => d.Semester).ToList();

                int row = 2;
                foreach (var discipline in sortedDisciplines)
                {
                    worksheet.Cells[row, 1].Value = discipline.Name;
                    worksheet.Cells[row, 2].Value = discipline.Semester;

                    var bestWorkload = bestVariant[discipline.UniqueName];
                    worksheet.Cells[row, 3].Value = bestWorkload;
                    worksheet.Cells[row, 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                    for (int i = 0; i < Math.Min(topVariants.Count, 10000); i++)
                    {
                        var variant = topVariants[i].variant;
                        var workload = variant[discipline.UniqueName];
                        worksheet.Cells[row, i + 4].Value = workload;
                    }
                    row++;
                }
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                await package.SaveAsAsync(new FileInfo(filePath));
            }
        }        private List<Discipline>? currentDisciplines;
        private List<(Dictionary<string, double> variant, double objective)>? currentTopVariants;
        private Dictionary<string, double>? currentBestVariant;

        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                    openFileDialog.Title = "Выберите файл Excel с данными";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            currentDisciplines = await ReadExcelDataAsync(openFileDialog.FileName);
                        }
                        catch (InvalidOperationException ex) when (ex.Message == "MultipleSheets")
                        {
                            using (var package = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                            {
                                var sheetNames = package.Workbook.Worksheets.Select(ws => ws.Name).ToArray();
                                using (var form = new Form())
                                {
                                    form.Text = "Выберите лист";
                                    form.Width = 300;
                                    form.Height = 150;
                                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                                    form.StartPosition = FormStartPosition.CenterParent;
                                    form.MaximizeBox = false;
                                    form.MinimizeBox = false;

                                    var label = new Label
                                    {
                                        Text = "Выберите лист для загрузки:",
                                        Location = new Point(10, 10),
                                        AutoSize = true
                                    };

                                    var comboBox = new ComboBox
                                    {
                                        Location = new Point(10, 40),
                                        Width = 260,
                                        DropDownStyle = ComboBoxStyle.DropDownList
                                    };
                                    comboBox.Items.AddRange(sheetNames);
                                    comboBox.SelectedIndex = 0;

                                    var button = new Button
                                    {
                                        Text = "OK",
                                        DialogResult = DialogResult.OK,
                                        Location = new Point(100, 70)
                                    };

                                    form.Controls.AddRange(new Control[] { label, comboBox, button });
                                    form.AcceptButton = button;                                    if (form.ShowDialog() == DialogResult.OK)
                                    {
                                        var selectedItem = comboBox.SelectedItem?.ToString();
                                        if (selectedItem != null)
                                        {
                                            currentDisciplines = await ReadExcelDataAsync(openFileDialog.FileName, selectedItem);
                                        }
                                    }
                                    else
                                    {
                                        return;
                                    }
                                }
                            }                        }

                        if (currentDisciplines == null)
                        {
                            MessageBox.Show("Не удалось загрузить данные из файла.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        var semDisc = PrepareSemDisc(currentDisciplines);
                        var validVariants = await GenerateValidVariantsAsync(currentDisciplines);

                        var scoredVariants = new ConcurrentBag<(Dictionary<string, double> variant, double objective)>();
                        Parallel.ForEach(validVariants, variant =>
                        {
                            double obj = CalculateObjectiveFast(semDisc, variant);
                            scoredVariants.Add((variant, obj));
                        });

                        currentTopVariants = scoredVariants
                            .OrderByDescending(x => x.objective)
                            .Take(10000)
                            .ToList();

                        if (currentTopVariants.Count > 0)
                        {
                            var best = currentTopVariants.First();
                            currentBestVariant = best.Item1;
                            Console.WriteLine("\nЛучший допустимый вариант:");
                            Console.WriteLine(string.Join(", ", best.Item1.Select(kv => $"{kv.Key}:{kv.Value}")));
                            label1.Text = $"Целевая функция: {best.Item2:F2}";

                            dataGridView1.Rows.Clear();

                            var sortedDisciplines = currentDisciplines.OrderBy(d => d.Semester).ToList();
  
                            foreach (var discipline in sortedDisciplines)
                            {
                                var workload = best.Item1[discipline.UniqueName];
                                dataGridView1.Rows.Add(
                                    discipline.Name,
                                    workload.ToString("F1"),
                                    discipline.SignificanceCoefficient.ToString("F2"),
                                    discipline.Semester
                                );
                            }
                        }
                        else
                        {
                            MessageBox.Show("Допустимых вариантов не найдено.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private async void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (currentDisciplines == null || currentTopVariants == null || currentBestVariant == null)
                {
                    MessageBox.Show("Сначала загрузите и обработайте данные!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "Сохранить результаты";
                    saveFileDialog.DefaultExt = "xlsx";
                    saveFileDialog.AddExtension = true;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        await SaveResultsToExcelAsync(saveFileDialog.FileName, currentDisciplines, currentBestVariant, currentTopVariants);
                        MessageBox.Show("Результаты успешно сохранены!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
