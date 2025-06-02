using ClosedXML.Excel;
using System.Data;

namespace TestOpt;

public partial class Form1 : Form
{
    private string currentFilePath = string.Empty;
    private XLWorkbook? workbook;
    private IXLWorksheet? worksheet;

    public Form1()
    {
        InitializeComponent();
    }

    private void btnOpenFile_Click(object sender, EventArgs e)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.Title = "Выберите Excel файл";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    currentFilePath = openFileDialog.FileName;
                    LoadExcelFile(currentFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии файла: {ex.Message}", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private void LoadExcelFile(string filePath)
    {
        try
        {
            // Закрываем предыдущий файл, если он был открыт
            workbook?.Dispose();

            // Открываем Excel файл
            workbook = new XLWorkbook(filePath);
            
            // Получаем первый лист
            worksheet = workbook.Worksheets.First();

            // Обновляем метку с именем файла
            lblFileName.Text = $"Файл: {Path.GetFileName(filePath)} | Лист: {worksheet.Name}";

            // Загружаем данные в DataGridView
            LoadDataToGrid();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при загрузке файла: {ex.Message}", "Ошибка", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }    private void LoadDataToGrid()
    {
        if (worksheet == null) return;

        try
        {
            // Получаем используемый диапазон
            var range = worksheet.RangeUsed();
            if (range == null)
            {
                dataGridView1.DataSource = null;
                MessageBox.Show("Лист пуст", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }            // Создаем DataTable для хранения данных
            DataTable dataTable = new DataTable();

            // Определяем количество строк и столбцов
            int rowCount = range.RowCount();
            int columnCount = range.ColumnCount();

            // Проверяем, есть ли заголовки (первая строка содержит текст)
            bool hasHeaders = false;
            for (int col = 1; col <= columnCount; col++)
            {
                var headerCell = worksheet.Cell(1, col);
                var headerValue = headerCell.Value.ToString().Trim();
                if (!string.IsNullOrEmpty(headerValue) && !double.TryParse(headerValue, out _))
                {
                    hasHeaders = true;
                    break;
                }
            }

            // Определяем диапазон данных
            int dataStartRow = hasHeaders ? 2 : 1;
            int dataRowCount = hasHeaders ? rowCount - 1 : rowCount;            // Сначала определяем типы данных для каждого столбца (анализируем только строки с данными)
            Type[] columnTypes = new Type[columnCount];
            for (int col = 1; col <= columnCount; col++)
            {
                columnTypes[col - 1] = DetermineColumnType(col, rowCount, hasHeaders);
            }

            // Показываем информацию о типах данных (для отладки)
            ShowColumnTypeInfo(columnCount, columnTypes, hasHeaders);

            // Добавляем столбцы в DataTable с правильными именами и типами
            for (int col = 1; col <= columnCount; col++)
            {
                string columnName;
                if (hasHeaders)
                {
                    // Используем заголовок из Excel
                    columnName = worksheet.Cell(1, col).Value.ToString().Trim();
                    if (string.IsNullOrEmpty(columnName))
                        columnName = GetExcelColumnName(col);
                }
                else
                {
                    // Используем стандартные имена столбцов (A, B, C...)
                    columnName = GetExcelColumnName(col);
                }
                dataTable.Columns.Add(columnName, columnTypes[col - 1]);
            }

            // Добавляем строки в DataTable (только данные, без заголовков)
            for (int row = dataStartRow; row <= rowCount; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= columnCount; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    dataRow[col - 1] = ConvertCellValue(cell, columnTypes[col - 1]);
                }
                dataTable.Rows.Add(dataRow);
            }            // Привязываем DataTable к DataGridView
            dataGridView1.DataSource = dataTable;

            // Настраиваем внешний вид DataGridView
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.ReadOnly = false;

            // Настраиваем форматирование для численных столбцов
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                var column = dataGridView1.Columns[i];
                if (columnTypes[i] == typeof(double))
                {
                    column.DefaultCellStyle.Format = "N2"; // 2 знака после запятой
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else if (columnTypes[i] == typeof(int))
                {
                    column.DefaultCellStyle.Format = "N0"; // Без знаков после запятой
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                }
            }

            // Показываем информацию о типах столбцов
            ShowColumnTypeInfo(columnCount, columnTypes, hasHeaders);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private string GetExcelColumnName(int columnNumber)
    {
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }    private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (worksheet == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;

        try
        {
            // Получаем новое значение из ячейки DataGridView
            var newValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            // Обновляем соответствующую ячейку в Excel
            if (newValue != null && newValue != DBNull.Value)
            {
                // Если это число, сохраняем как число, иначе как строку
                if (newValue is double || newValue is int || newValue is decimal)
                {
                    worksheet.Cell(e.RowIndex + 1, e.ColumnIndex + 1).Value = Convert.ToDouble(newValue);
                }
                else
                {
                    worksheet.Cell(e.RowIndex + 1, e.ColumnIndex + 1).Value = newValue.ToString();
                }
            }
            else
            {
                worksheet.Cell(e.RowIndex + 1, e.ColumnIndex + 1).Value = "";
            }

            // Сохраняем изменения в файл
            SaveChangesToFile();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при сохранении изменений: {ex.Message}", "Ошибка", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SaveChangesToFile()
    {
        if (workbook == null || string.IsNullOrEmpty(currentFilePath)) return;

        try
        {
            workbook.Save();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        // Закрываем workbook при закрытии формы
        workbook?.Dispose();
        base.OnFormClosed(e);
    }    private Type DetermineColumnType(int columnIndex, int rowCount, bool hasHeaders = false)
    {
        if (worksheet == null) return typeof(string);

        // Определяем диапазон для анализа
        int startRow = hasHeaders ? 2 : 1; // Если есть заголовки, начинаем со 2 строки
        int endRow = Math.Min(rowCount, startRow + 9); // Анализируем до 10 строк
        
        // Проверяем заголовок столбца для подсказок (только если есть заголовки)
        bool isLikelyNumeric = false;
        if (hasHeaders)
        {
            var headerCell = worksheet.Cell(1, columnIndex);
            var headerText = headerCell.Value.ToString().ToLower();
            
            isLikelyNumeric = headerText.Contains("трудоемкость") || headerText.Contains("коэф") || 
                headerText.Contains("семестр") || headerText.Contains("номер") ||
                headerText.Contains("№") || headerText.Contains("значимость");
        }
        
        int numericCount = 0;
        int integerCount = 0;
        int totalChecked = 0;

        for (int row = startRow; row <= endRow; row++)
        {
            var cell = worksheet.Cell(row, columnIndex);
            var cellValue = cell.Value.ToString().Trim();
            
            if (string.IsNullOrEmpty(cellValue)) continue;
            
            totalChecked++;

            // Проверяем, является ли значение числом
            if (double.TryParse(cellValue, out double doubleValue))
            {
                numericCount++;
                // Проверяем, является ли число целым
                if (doubleValue == Math.Floor(doubleValue))
                {
                    integerCount++;
                }
            }
        }

        // Если заголовок указывает на числовые данные или большинство значений числовые
        if (isLikelyNumeric && numericCount > 0 || numericCount > totalChecked / 2)
        {
            // Если все числа целые, используем int, иначе double
            return integerCount == numericCount ? typeof(int) : typeof(double);
        }

        return typeof(string);
    }

    private object ConvertCellValue(IXLCell cell, Type targetType)
    {
        try
        {
            var cellValue = cell.Value.ToString().Trim();
            
            if (string.IsNullOrEmpty(cellValue))
            {
                if (targetType == typeof(string)) return string.Empty;
                if (targetType == typeof(int)) return 0;
                if (targetType == typeof(double)) return 0.0;
                return DBNull.Value;
            }

            if (targetType == typeof(string))
            {
                return cellValue;
            }
            else if (targetType == typeof(int))
            {
                if (int.TryParse(cellValue, out int intValue))
                    return intValue;
                // Если не удается парсить как int, пробуем double и округляем
                if (double.TryParse(cellValue, out double doubleValue))
                    return (int)Math.Round(doubleValue);
                return 0;
            }
            else if (targetType == typeof(double))
            {
                if (double.TryParse(cellValue, out double doubleValue))
                    return doubleValue;
                return 0.0;
            }

            return cellValue;
        }
        catch
        {
            // В случае ошибки возвращаем значение по умолчанию для типа
            if (targetType == typeof(string)) return string.Empty;
            if (targetType == typeof(int)) return 0;
            if (targetType == typeof(double)) return 0.0;
            return DBNull.Value;
        }
    }    private void ShowColumnTypeInfo(int columnCount, Type[] columnTypes, bool hasHeaders)
    {
        if (worksheet == null) return;
        
        string info = $"Обнаружено заголовков: {(hasHeaders ? "Да" : "Нет")}\n\n";
        info += "Типы столбцов:\n";
        
        for (int i = 0; i < columnCount; i++)
        {
            string columnName;
            if (hasHeaders)
            {
                columnName = worksheet.Cell(1, i + 1).Value.ToString().Trim();
            }
            else
            {
                columnName = GetExcelColumnName(i + 1);
            }
            
            string typeName = columnTypes[i] == typeof(int) ? "Целое число" :
                             columnTypes[i] == typeof(double) ? "Дробное число" : "Текст";
            
            info += $"{columnName}: {typeName}\n";
        }
        
        MessageBox.Show(info, "Информация о типах данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
