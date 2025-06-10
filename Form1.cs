using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Collections.Concurrent;
using Optimizations; // Добавлено для доступа к Discipline, ExcelService, OptimizationService


namespace Optimizations
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
            statusLabel.Text = "Готов к работе";
            progressBar.Visible = false;
            
            button3.Enabled = false;
            
            this.Icon = SystemIcons.Application;
        }

        private List<Discipline>? currentDisciplines;
        private List<(Dictionary<string, double> variant, double objective)>? currentTopVariants;
        private Dictionary<string, double>? currentBestVariant;

        private async void button1_Click(object sender, EventArgs e)
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
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        statusLabel.Text = "Загрузка данных из Excel файла...";
                        button1.Enabled = false;
                        button3.Enabled = false;

                        try
                        {
                            currentDisciplines = await ExcelService.ReadExcelDataAsync(openFileDialog.FileName);
                        }
                        catch (InvalidOperationException ex) when (ex.Message == "MultipleSheets")
                        {
                            statusLabel.Text = "Выбор листа Excel...";
                            using (var excelPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(openFileDialog.FileName)))
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
                                    okButton.FlatAppearance.BorderSize = 0;
                                    sheetSelectionForm.Controls.AddRange(new Control[] { instructionLabel, sheetComboBox, okButton });
                                    sheetSelectionForm.AcceptButton = okButton;

                                    if (sheetSelectionForm.ShowDialog() == DialogResult.OK)
                                    {
                                        var selectedSheetName = sheetComboBox.SelectedItem?.ToString();
                                        if (selectedSheetName != null)
                                        {
                                            statusLabel.Text = $"Загрузка данных из листа '{selectedSheetName}'...";
                                            currentDisciplines = await ExcelService.ReadExcelDataAsync(openFileDialog.FileName, selectedSheetName);
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

                        if (currentDisciplines == null || currentDisciplines.Count == 0)
                        {
                            MessageBox.Show("Не удалось загрузить данные из файла или файл не содержит дисциплин.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            progressBar.Visible = false;
                            statusLabel.Text = "Ошибка загрузки данных";
                            button1.Enabled = true;
                            return;
                        }

                        HashSet<string> excludedDisciplineNames = new HashSet<string>();
                        using (var excludeForm = new ExcludeDisciplinesForm(currentDisciplines))
                        {
                            if (excludeForm.ShowDialog() == DialogResult.OK)
                            {
                                excludedDisciplineNames = excludeForm.ExcludedDisciplineNames;
                            }
                            else
                            {
                                statusLabel.Text = "Выбор исключаемых дисциплин отменен. Оптимизация без исключений.";
                            }
                        }

                        statusLabel.Text = "Генерация вариантов оптимизации...";
                        
                        var targetSums = GetTargetSumsFromTextBoxes();
                        if (targetSums == null)
                        {
                            progressBar.Visible = false;
                            statusLabel.Text = "Ошибка в целевых суммах";
                            button1.Enabled = true;
                            return;
                        }
                        
                        if (!ValidateTargetSums(currentDisciplines, targetSums))
                        {
                            progressBar.Visible = false;
                            statusLabel.Text = "Оптимизация отменена пользователем";
                            button1.Enabled = true;
                            return;
                        }
                        
                        var semesterDisciplinesArray = OptimizationService.PrepareSemDisc(currentDisciplines);
                        var validVariants = await OptimizationService.GenerateValidVariantsAsync(currentDisciplines, excludedDisciplineNames, targetSums);
                        statusLabel.Text = "Расчет целевой функции...";
                        var scoredVariants = new ConcurrentBag<(Dictionary<string, double> variant, double objective)>();
                        Parallel.ForEach(validVariants, workloadVariant =>
                        {
                            double objectiveValue = OptimizationService.CalculateObjectiveFast(semesterDisciplinesArray, workloadVariant);
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
                            currentBestVariant = bestVariantTuple.Item1;

                            Console.WriteLine("\nЛучший допустимый вариант:");
                            Console.WriteLine(string.Join(", ", bestVariantTuple.Item1.Select(keyValue => $"{keyValue.Key}:{keyValue.Value}")));
                            
                            statusLabel.Text = "Заполнение таблицы результатов...";
                            
                            statusLabel.Text = "Готово! Оптимизация завершена успешно";
                            button3.Enabled = true;
                        }
                        else
                        {
                            MessageBox.Show("Допустимых вариантов не найдено.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            statusLabel.Text = "Допустимых вариантов не найдено";
                        }
                        
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

                using (SaveFileDialog saveResultsDialog = new SaveFileDialog())
                {
                    saveResultsDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                    saveResultsDialog.Title = "Сохранить результаты";
                    saveResultsDialog.DefaultExt = "xlsx";
                    saveResultsDialog.AddExtension = true;

                    if (saveResultsDialog.ShowDialog() == DialogResult.OK)
                    {
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        statusLabel.Text = "Сохранение результатов в Excel файл...";
                        button3.Enabled = false;
                        button1.Enabled = false;
                        
                        int variantsToSave = Math.Min(currentTopVariants.Count, 50);
                        await ExcelService.SaveResultsToExcelAsync(saveResultsDialog.FileName, currentDisciplines, currentBestVariant, currentTopVariants);
                        
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
        }

        private List<(int Sem1, int Sem2, double TargetSum)>? GetTargetSumsFromTextBoxes()
        {
            try
            {
                var sum12 = double.Parse(textBoxSem12.Text.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                var sum34 = double.Parse(textBoxSem34.Text.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                var sum56 = double.Parse(textBoxSem56.Text.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                var sum78 = double.Parse(textBoxSem78.Text.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);

                // Валидация разумности целевых сумм
                if (sum12 <= 0 || sum34 <= 0 || sum56 <= 0 || sum78 <= 0)
                {
                    MessageBox.Show("Целевые суммы должны быть положительными числами.", 
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }

                if (sum12 > 200 || sum34 > 200 || sum56 > 200 || sum78 > 200)
                {
                    var result = MessageBox.Show("Одна или несколько целевых сумм кажутся очень большими (>200). Продолжить?", 
                        "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.No)
                        return null;
                }

                return new List<(int Sem1, int Sem2, double TargetSum)>
                {
                    (1, 2, sum12),
                    (3, 4, sum34),
                    (5, 6, sum56),
                    (7, 8, sum78)
                };
            }
            catch (FormatException)
            {
                MessageBox.Show("Некорректные значения в полях целевых сумм. Используйте числовые значения.", 
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        // Добавим метод для проверки совместимости целевых сумм с данными
        private bool ValidateTargetSums(List<Discipline> disciplines, List<(int Sem1, int Sem2, double TargetSum)> targetSums)
        {
            var warnings = new List<string>();
            
            foreach (var (sem1, sem2, targetSum) in targetSums)
            {
                var blockDisciplines = disciplines.Where(d => d.Semester == sem1 || d.Semester == sem2).ToList();
                if (blockDisciplines.Count == 0)
                {
                    warnings.Add($"В семестрах {sem1}-{sem2} нет дисциплин");
                    continue;
                }
                
                var minPossibleSum = blockDisciplines.Sum(d => d.MinWorkload);
                var maxPossibleSum = blockDisciplines.Sum(d => d.MaxWorkload);
                
                if (targetSum < minPossibleSum || targetSum > maxPossibleSum)
                {
                    warnings.Add($"Семестры {sem1}-{sem2}: целевая сумма {targetSum} недостижима (возможно: {minPossibleSum:F1}-{maxPossibleSum:F1})");
                }
            }
            
            if (warnings.Count > 0)
            {
                var warningMessage = "Обнаружены потенциальные проблемы:\n\n" + string.Join("\n", warnings) + "\n\nПродолжить оптимизацию?";
                var result = MessageBox.Show(warningMessage, "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                return result == DialogResult.Yes;
            }
            
            return true;
        }
    }
}
