﻿using ClosedXML.Excel;
using System;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            string fileName = "na_avtomat.xlsx";
            Console.WriteLine($"Файл: {fileName}");
            Dictionary<string, float[]> dictFromExcel =  new Dictionary<string, float[]>(); 

            using (var workbook = new XLWorkbook(fileName))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    Console.WriteLine($"\nСтраница: {worksheet.Name}");
                    
                    var range = worksheet.RangeUsed();
                    if (range != null)
                    {
                        foreach (var row in range.Rows())
                        {
                            var cells = row.Cells().ToList();
                            if (cells.Count > 0)
                            {
                                string key = cells[0].Value.ToString();
                                float[] values = cells.Skip(1)
                                    .Select(c => {
                                        if (c.IsEmpty()) return 0f;
                                        string valStr = c.Value.ToString().Trim().Replace("\"", "");
                                        if (string.IsNullOrWhiteSpace(valStr)) return 0f;
                                        
                                        // Пробуем разные форматы чисел
                                        if (float.TryParse(valStr, System.Globalization.NumberStyles.Any, 
                                            System.Globalization.CultureInfo.InvariantCulture, out float num))
                                            return num;
                                            
                                        if (float.TryParse(valStr, System.Globalization.NumberStyles.Any, 
                                            System.Globalization.CultureInfo.GetCultureInfo("ru-RU"), out num))
                                            return num;
                                            
                                        return 0f;
                                    })
                                    .ToArray();
                                
                                dictFromExcel[key] = values;
                                
                                Console.WriteLine($"Добавлено в словарь: {key} => [{string.Join("; ", values)}]");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Пустая страница");
                    }
                }
            }

            Console.WriteLine("\nСодержимое словаря:");
            foreach (var item in dictFromExcel)
            {
                Console.WriteLine($"{item.Key}: [{string.Join("; ", item.Value)}]");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }
}
