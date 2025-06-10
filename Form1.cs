using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml; 
using System.Collections.Concurrent;


namespace Optimizations
{
    public partial class Form1 : Form
    {
        private List<Discipline>? currentDisciplines;
        // Changed type: no longer stores objective
        private List<Dictionary<string, double>>? currentTopVariants; 
        private Dictionary<string, double>? currentBestVariant;
        private HashSet<string> excludedDisciplineNames = new HashSet<string>();

        public Form1()
        {
            InitializeComponent();
            
            // Button text is set in designer, no need to set here unless dynamic
            // if (this.button1 != null) this.button1.Text = "📁 Загрузить файл"; // Already set in designer
            // if (this.button3 != null) this.button3.Text = "💾 Сохранить"; // Already set in designer
            // if (this.button2 != null) this.button2.Text = "❌ Закрыть"; // Already set in designer

            // Initial status
            statusLabel.Text = "✅ Готов к работе";
            progressBar.Visible = false;
        }

        private List<(int Sem1, int Sem2, double TargetSum)> GetTargetSemesterSums()
        {
            var sums = new List<(int Sem1, int Sem2, double TargetSum)>();
            try
            {
                sums.Add((1, 2, double.Parse(textBoxSem12.Text.Replace(",", "."), CultureInfo.InvariantCulture)));
                sums.Add((3, 4, double.Parse(textBoxSem34.Text.Replace(",", "."), CultureInfo.InvariantCulture)));
                sums.Add((5, 6, double.Parse(textBoxSem56.Text.Replace(",", "."), CultureInfo.InvariantCulture)));
                sums.Add((7, 8, double.Parse(textBoxSem78.Text.Replace(",", "."), CultureInfo.InvariantCulture)));
            }
            catch (FormatException ex)
            {
                MessageBox.Show($"Ошибка в формате целевых сумм семестров: {ex.Message}", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw; // Re-throw to stop processing
            }
            return sums;
        }

        private async void button1_Click(object? sender, EventArgs e) 
        {
            try
            {
                // Step 1: Load Excel File
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                    openFileDialog.Title = "Выберите файл Excel с данными";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;
                        string? selectedSheetName = null;

                        statusLabel.Text = "⏳ Чтение файла Excel...";
                        infoLabel.Text = "Идет загрузка данных из файла...";
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        Application.DoEvents();

                        try
                        {
                            currentDisciplines = await ExcelService.ReadExcelDataAsync(filePath);
                        }
                        catch (InvalidOperationException ex) when (ex.Message == "MultipleSheets")
                        {
                            using (var package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                var sheetNames = package.Workbook.Worksheets.Select(ws => ws.Name).ToArray();
                                using (var sheetForm = new Form())
                                {
                                    sheetForm.Text = "Выберите лист";
                                    sheetForm.Width = 300;
                                    sheetForm.Height = 150;
                                    sheetForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                                    sheetForm.StartPosition = FormStartPosition.CenterParent;
                                    sheetForm.MaximizeBox = false;
                                    sheetForm.MinimizeBox = false;

                                    var label = new Label { Text = "В файле несколько листов. Выберите один:", Location = new Point(10, 10), AutoSize = true };
                                    var comboBox = new ComboBox { Location = new Point(10, 40), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
                                    comboBox.Items.AddRange(sheetNames);
                                    if (comboBox.Items.Count > 0) comboBox.SelectedIndex = 0;
                                    var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(100, 70) };

                                    sheetForm.Controls.AddRange(new Control[] { label, comboBox, okButton });
                                    sheetForm.AcceptButton = okButton;


                                    if (sheetForm.ShowDialog(this) == DialogResult.OK && comboBox.SelectedItem != null)
                                    {
                                        selectedSheetName = comboBox.SelectedItem.ToString();
                                        currentDisciplines = await ExcelService.ReadExcelDataAsync(filePath, selectedSheetName);
                                    }
                                    else
                                    {
                                        infoLabel.Text = "Загрузите Excel файл для оптимизации расписания";
                                        statusLabel.Text = "⚠️ Чтение файла отменено.";
                                        progressBar.Visible = false;
                                        return; 
                                    }
                                }
                            }
                        }
                        
                        if (currentDisciplines == null || currentDisciplines.Count == 0)
                        {
                            MessageBox.Show("Не удалось загрузить дисциплины из файла.", "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            infoLabel.Text = "Загрузите Excel файл для оптимизации расписания";
                            statusLabel.Text = "⚠️ Ошибка загрузки данных.";
                            progressBar.Visible = false;
                            return;
                        }

                        statusLabel.Text = "✅ Данные загружены. Укажите исключения.";
                        infoLabel.Text = "Данные из Excel успешно загружены.";
                        progressBar.Visible = false;
                        Application.DoEvents();

                        // Step 2: Select Exclusions
                        using (var excludeForm = new ExcludeDisciplinesForm(currentDisciplines, excludedDisciplineNames))
                        {
                            if (excludeForm.ShowDialog(this) == DialogResult.OK)
                            {
                                excludedDisciplineNames = excludeForm.ExcludedDisciplineNames;
                                statusLabel.Text = "⏳ Подготовка к оптимизации...";
                                infoLabel.Text = "Исключения приняты. Введите целевые суммы и запустите оптимизацию.";
                                // At this point, the user has confirmed exclusions.
                                // Now, get target sums and proceed with optimization.
                            }
                            else
                            {
                                infoLabel.Text = "Загрузите Excel файл для оптимизации расписания";
                                statusLabel.Text = "⚠️ Выбор исключений отменен. Оптимизация не запущена.";
                                progressBar.Visible = false;
                                return; 
                            }
                        }

                        // Step 3: Get Target Sums (already part of the flow, but now after exclusions)
                        List<(int Sem1, int Sem2, double TargetSum)> targetSemesterSums;
                        try
                        {
                            targetSemesterSums = GetTargetSemesterSums();
                        }
                        catch
                        {
                            statusLabel.Text = "⚠️ Ошибка в целевых суммах. Оптимизация не запущена.";
                            progressBar.Visible = false;
                            return;
                        }
                        
                        // Step 4: Perform Optimization
                        infoLabel.Text = "Идет оптимизация расписания...";
                        statusLabel.Text = "⏳ Идет оптимизация...";
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        Application.DoEvents(); 

                        var validVariants = await OptimizationService.GenerateValidVariantsAsync(currentDisciplines, excludedDisciplineNames, targetSemesterSums);

                        progressBar.Style = ProgressBarStyle.Continuous;
                        progressBar.Value = 0; 

                        if (validVariants == null || validVariants.Count == 0)
                        {
                             // Removed: MessageBox.Show("Не найдено допустимых вариантов после генерации.", "Нет вариантов", MessageBoxButtons.OK, MessageBoxIcon.Information);
                             infoLabel.Text = "Допустимых вариантов не найдено. Сохранение исходных данных...";
                             statusLabel.Text = "ℹ️ Допустимых вариантов не найдено. Сохранение исходных данных...";
                             // progressBar.Visible = false; // Keep progress bar for saving operation

                             // New logic: Directly save initial data if no variants are found
                             if (currentDisciplines != null && currentDisciplines.Count > 0)
                             {
                                 // Removed: DialogResult dialogResult = MessageBox.Show(...)
                                 using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                                 {
                                     saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                                     saveFileDialog.Title = "Сохранить исходные данные (вариантов не найдено)";
                                     saveFileDialog.DefaultExt = "xlsx";
                                     saveFileDialog.AddExtension = true;
                                     saveFileDialog.FileName = "Data.xlsx"; // Changed filename to be more specific

                                     if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                     {
                                         try
                                         {
                                             statusLabel.Text = "⏳ Сохранение исходных данных...";
                                             infoLabel.Text = "Идет сохранение исходных данных в Excel...";
                                             progressBar.Visible = true;
                                             progressBar.Style = ProgressBarStyle.Marquee;
                                             Application.DoEvents();

                                             await ExcelService.SaveInitialDataAsSingleVariantAsync(saveFileDialog.FileName, currentDisciplines);
                                             
                                             MessageBox.Show("Исходные данные (как единственный вариант) успешно сохранены, так как других вариантов не найдено.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                             statusLabel.Text = "✅ Исходные данные сохранены.";
                                             infoLabel.Text = "Исходные данные сохранены. Готов к новой задаче.";
                                         }
                                         catch (Exception exSave)
                                         {
                                             MessageBox.Show($"Произошла ошибка при сохранении исходных данных: {exSave.Message}", "Ошибка сохранения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                             statusLabel.Text = "❌ Ошибка сохранения исходных данных.";
                                             infoLabel.Text = "Произошла ошибка при сохранении исходных данных.";
                                         }
                                         finally
                                         {
                                             progressBar.Visible = false;
                                         }
                                     }
                                     else
                                     {
                                         statusLabel.Text = "⚠️ Сохранение исходных данных отменено пользователем.";
                                         infoLabel.Text = "Оптимизация завершена, вариантов не найдено. Сохранение отменено.";
                                         progressBar.Visible = false;
                                     }
                                 }
                             }
                             else // currentDisciplines is null or empty, should not happen if load was successful
                             {
                                 MessageBox.Show("Исходные данные для сохранения отсутствуют.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                 statusLabel.Text = "⚠️ Ошибка: нет исходных данных для сохранения.";
                                 progressBar.Visible = false;
                             }
                             return;
                        }

                        currentTopVariants = validVariants; 

                        if (currentTopVariants.Count > 0)
                        {
                            currentBestVariant = currentTopVariants.FirstOrDefault(); 
                            statusLabel.Text = $"🏆 Оптимизация завершена! Найдено {currentTopVariants.Count} вариантов.";
                            infoLabel.Text = "Оптимизация завершена. Результаты готовы к сохранению.";
                        }
                        else
                        {
                            MessageBox.Show("Допустимых вариантов не найдено.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            infoLabel.Text = "Оптимизация завершена.";
                            statusLabel.Text = "ℹ️ Допустимых вариантов не найдено.";
                        }
                        progressBar.Visible = false;
                    }
                    else // User cancelled OpenFileDialog
                    {
                        infoLabel.Text = "Загрузите Excel файл для оптимизации расписания";
                        statusLabel.Text = "✅ Готов к работе";
                        progressBar.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}\n{ex.StackTrace}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                infoLabel.Text = "Загрузите Excel файл для оптимизации расписания";
                statusLabel.Text = "❌ Произошла ошибка.";
                progressBar.Visible = false;
            }
        }

        private async void button3_Click_1(object? sender, EventArgs e) 
        {
            if (currentDisciplines == null || currentTopVariants == null || currentBestVariant == null) // currentBestVariant could be null if currentTopVariants is empty
            {
                MessageBox.Show("Сначала загрузите и обработайте данные! Валидных вариантов не найдено.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                statusLabel.Text = "⚠️ Нет данных для сохранения.";
                return;
            }
            if (currentTopVariants.Count == 0 || currentBestVariant == null)
            {
                 MessageBox.Show("Нет вариантов для сохранения.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                statusLabel.Text = "⚠️ Нет данных для сохранения.";
                return;
            }

            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "Сохранить результаты";
                    saveFileDialog.DefaultExt = "xlsx";
                    saveFileDialog.AddExtension = true;
                    saveFileDialog.FileName = "Optimization_Results.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        statusLabel.Text = "⏳ Сохранение результатов...";
                        infoLabel.Text = "Идет сохранение результатов в Excel...";
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;
                        Application.DoEvents();

                        // Pass currentTopVariants directly
                        await ExcelService.SaveResultsToExcelAsync(saveFileDialog.FileName, currentDisciplines, currentBestVariant, currentTopVariants); 
                        
                        MessageBox.Show("Результаты успешно сохранены!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        statusLabel.Text = "✅ Результаты сохранены.";
                        infoLabel.Text = "Результаты сохранены. Готов к новой задаче.";
                        progressBar.Visible = false;
                    }
                    else
                    {
                        statusLabel.Text = "⚠️ Сохранение отменено.";
                        progressBar.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при сохранении: {ex.Message}", "Ошибка сохранения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "❌ Ошибка сохранения.";
                infoLabel.Text = "Произошла ошибка при сохранении.";
                progressBar.Visible = false;
            }
        }

        // private void exitButton_Click(object? sender, EventArgs e) // REMOVED
        // {
        //     this.Close();
        // }

        private void button2_Click(object? sender, EventArgs e) // New handler for "Закрыть" button
        {
            this.Close();
        }
    }
}
