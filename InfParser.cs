using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace K450DataExporter
{
    public class InfParser
    {
        // Основные параметры из GENERAL
        public string FileDati { get; private set; } = "";
        public string FileSettimanale { get; private set; } = "";

        // Резервная копия из ODR
        public string NomeFileCopia { get; private set; } = "";

        // Список полей PROCESS_TABLE (только валидные)
        public List<ProcessField> ProcessFields { get; private set; } = new List<ProcessField>();

        // Список полей PRODUCT_TABLE (все поля)
        public List<ProductField> ProductFields { get; private set; } = new List<ProductField>();

        // GRID настройки
        public List<string> GridProcessFields { get; private set; } = new List<string>();
        public List<string> GridProductFields { get; private set; } = new List<string>();

        // PROCESS_REPORT1 (для отображения отчета)
        public List<ReportField> ProcessReportFields { get; private set; } = new List<ReportField>();

        // PRODUCT_REPORT1 (для отображения отчета)
        public List<ReportField> ProductReportFields { get; private set; } = new List<ReportField>();

        // MINMAX1 (заглушка)
        public List<string> MinMaxFields { get; private set; } = new List<string>();

        // Имена таблиц в базе данных (жестко прописаны)
        public const string ProcessTableName = "Process_DataE";
        public const string ProductTableName = "Production_DataE";

        public InfParser(string infPath)
        {
            if (!File.Exists(infPath))
                throw new FileNotFoundException($"INF-файл не найден: {infPath}");

            ParseInfFile(infPath);
        }

        private void ParseInfFile(string infPath)
        {
            string[] lines = File.ReadAllLines(infPath, Encoding.Default);
            string currentSection = "";

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();

                if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith(";"))
                    continue;

                if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                {
                    currentSection = trimmedLine.TrimStart('[').TrimEnd(']').ToUpper();
                    continue;
                }

                if (trimmedLine.Contains("="))
                {
                    string[] parts = trimmedLine.Split(new char[] { '=' }, 2);
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();

                    switch (currentSection)
                    {
                        case "GENERAL":
                            if (key == "FileDati") FileDati = value;
                            if (key == "FileSettimanale") FileSettimanale = value;
                            break;

                        case "ODR":
                            if (key == "NomeFileCopia") NomeFileCopia = value;
                            break;

                        case "PROCESS_TABLE":
                            if (key.StartsWith("Campo"))
                            {
                                var field = new ProcessField(value);
                                // Проверяем валидность: PlcAddress не null и не пуст, Description не null
                                if (field.IsValid)
                                {
                                    ProcessFields.Add(field);
                                }
                            }
                            break;

                        case "PRODUCT_TABLE":
                            if (key.StartsWith("Campo"))
                            {
                                ProductFields.Add(new ProductField(value));
                            }
                            break;

                        case "GRID":
                            if (key == "Process")
                            {
                                GridProcessFields.AddRange(value.Split(';'));
                            }
                            else if (key == "Product")
                            {
                                GridProductFields.AddRange(value.Split(';'));
                            }
                            break;

                        case "PROCESS_REPORT1":
                            if (key.StartsWith("Campo"))
                            {
                                ProcessReportFields.Add(new ReportField(value));
                            }
                            break;

                        case "PRODUCT_REPORT1":
                            if (key.StartsWith("Campo"))
                            {
                                ProductReportFields.Add(new ReportField(value));
                            }
                            break;

                        case "MINMAX1":
                            if (key == "CampiMinMax")
                            {
                                MinMaxFields.AddRange(value.Split(';'));
                            }
                            break;
                    }
                }
            }
        }

        public string GetAvailableDatabasePath(string baseDirectory)
        {
            string primaryPath = Path.Combine(baseDirectory, FileDati);
            if (File.Exists(primaryPath))
                return primaryPath;

            if (!string.IsNullOrEmpty(NomeFileCopia))
            {
                string copyPath = Path.Combine(baseDirectory, NomeFileCopia);
                if (File.Exists(copyPath))
                    return copyPath;
            }

            return primaryPath;
        }

        /// <summary>
        /// Получить список полей для ODR таблицы из GRID
        /// </summary>

        public List<GridFieldInfo> GetGridProcessFields()
        {
            var result = new List<GridFieldInfo>();

            foreach (string fieldName in GridProcessFields)
            {
                var processField = ProcessFields.FirstOrDefault(f => f.Name == fieldName);

                if (processField != null)
                {
                    result.Add(new GridFieldInfo
                    {
                        FieldName = processField.Name,        // PRC001 (имя колонки в базе!)
                        Alias = processField.Description,     // Cycle time (человеческое название)
                        Unit = processField.Unit,
                        ArchiveType = processField.ArchiveType
                    });
                }
                else
                {
                    // Если поле не найдено, пробуем использовать как есть
                    result.Add(new GridFieldInfo
                    {
                        FieldName = fieldName,
                        Alias = fieldName,
                        Unit = "",
                        ArchiveType = ""
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Получить список полей для PDR таблицы из GRID
        /// </summary>
        public List<GridFieldInfo> GetGridProductFields()
        {
            var result = new List<GridFieldInfo>();

            foreach (string fieldName in GridProductFields)
            {
                var productField = ProductFields.FirstOrDefault(f => f.Name == fieldName);

                if (productField != null)
                {
                    result.Add(new GridFieldInfo
                    {
                        FieldName = fieldName,
                        Alias = productField.Description,
                        Unit = productField.Unit,
                        ArchiveType = productField.ArchiveType
                    });
                }
                else
                {
                    result.Add(new GridFieldInfo
                    {
                        FieldName = fieldName,
                        Alias = fieldName,
                        Unit = "",
                        ArchiveType = ""
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Построить SELECT запрос для PROCESS_TABLE
        /// </summary>
        public string BuildProcessSelectQuery()
        {
            var gridFields = GetGridProcessFields();
            var selectParts = new List<string> { "[Data]", "[Ora]", "[Cy] AS Cycle" };

            foreach (var field in gridFields)
            {
                string cleanName = field.FieldName.Replace("[", "").Replace("]", "");
                string alias = field.Alias;
                if (!string.IsNullOrEmpty(field.Unit) && field.Unit != "#" && field.Unit != "null")
                {
                    alias = $"{field.Alias} ({field.Unit})";
                }
                selectParts.Add($"[{cleanName}] AS [{alias}]");
            }

            return string.Join(", ", selectParts);
        }
    }

    /// <summary>
    /// Структура поля PROCESS_TABLE
    /// Валидным считается поле, у которого PlcAddress не null и Description не null
    /// </summary>
    public class ProcessField
    {
        public string Name { get; set; }          // PRC001 или [CBAP]
        public string Type { get; set; }           // 7
        public string Decimals { get; set; }       // 3
        public string Expression { get; set; }     // null
        public string PlcAddress { get; set; }     // DB121,REAL12 (ключевой параметр 5)
        public string Scale { get; set; }          // 1
        public string Offset { get; set; }         // 0
        public string ShortName { get; set; }      // TCY или [CBAP]
        public string Description { get; set; }    // Cycle time или Combustion air pressure (ключевой параметр 9)
        public string MinValue { get; set; }       // 0
        public string MaxValue { get; set; }       // 700
        public string Unit { get; set; }           // #
        public string AlarmMin { get; set; }       // 0
        public string AlarmMax { get; set; }       // 9999
        public string ArchiveType { get; set; }    // S
        public string Alias { get; set; }          // Cycle time

        /// <summary>
        /// Очищенное имя поля (без квадратных скобок)
        /// </summary>
        public string CleanName => Name?.Replace("[", "").Replace("]", "");

        /// <summary>
        /// Поле считается валидным, если PlcAddress не пустой и Description не пустой
        /// </summary>
        public bool IsValid => !string.IsNullOrEmpty(PlcAddress) && PlcAddress != "null" &&
                               !string.IsNullOrEmpty(Description) && Description != "null";

        public ProcessField(string rawData)
        {
            string[] parts = rawData.Split(';');
            if (parts.Length >= 19)
            {
                Name = parts[0];
                Type = parts[1];
                Decimals = parts[2];
                Expression = parts[3];
                PlcAddress = parts[4];
                Scale = parts[5];
                Offset = parts[6];
                ShortName = parts[7];
                Description = parts[8];
                MinValue = parts[9];
                MaxValue = parts[10];
                Unit = parts[11];
                AlarmMin = parts[12];
                AlarmMax = parts[13];
                ArchiveType = parts[14];
                Alias = parts[15];
            }
            else
            {
                Name = parts[0];
                Description = parts.Length > 8 ? parts[8] : Name;
                Alias = Description;
            }
        }
    }

    /// <summary>
    /// Структура поля PRODUCT_TABLE
    /// </summary>
    public class ProductField
    {
        public string Name { get; set; }          // PRD001
        public string Type { get; set; }           // 7
        public string Decimals { get; set; }       // 3
        public string Formula { get; set; }        // SUM(PRC002)/1000 (формула!)
        public string Scale { get; set; }          // 1
        public string Offset { get; set; }         // 0
        public string ShortName { get; set; }      // Lime
        public string Description { get; set; }    // Quiklime produced
        public string MinValue { get; set; }       // 0
        public string MaxValue { get; set; }       // 1000
        public string Unit { get; set; }           // t
        public string AlarmMin { get; set; }       // 0
        public string AlarmMax { get; set; }       // 50
        public string ArchiveType { get; set; }    // S
        public string Alias { get; set; }          // Quiklime produced

        public ProductField(string rawData)
        {
            string[] parts = rawData.Split(';');
            if (parts.Length >= 17)
            {
                Name = parts[0];
                Type = parts[1];
                Decimals = parts[2];
                Formula = parts[3];
                Scale = parts[4];
                Offset = parts[5];
                ShortName = parts[6];
                Description = parts[7];
                MinValue = parts[8];
                MaxValue = parts[9];
                Unit = parts[10];
                AlarmMin = parts[11];
                AlarmMax = parts[12];
                ArchiveType = parts[13];
                Alias = parts[14];
            }
            else
            {
                Name = parts[0];
                Alias = parts.Length > 7 ? parts[7] : Name;
                Description = Alias;
            }
        }
    }

    /// <summary>
    /// Структура поля для REPORT секций
    /// </summary>
        public class ReportField
        {
            public string Name { get; set; }
            public string Unit { get; set; }
            public string DefaultValue { get; set; }
            public string Formula { get; set; }
            public string Alias { get; set; }

            // ПУСТОЙ КОНСТРУКТОР (ДОБАВИТЬ!)
            public ReportField() { }

            public ReportField(string rawData)
            {
                string[] parts = rawData.Split(';');
                if (parts.Length >= 5)
                {
                    Name = parts[0];
                    Unit = parts.Length > 1 ? parts[1] : "";
                    DefaultValue = parts.Length > 2 ? parts[2] : "0";
                    Formula = parts.Length > 3 ? parts[3] : "";
                    Alias = parts.Length > 4 ? parts[4] : Name;
                }
                else
                {
                    Name = rawData;
                    Alias = Name;
                }
            }
        }

        /// <summary>
        /// Информация о поле для отображения в сетке
        /// </summary>
        public class GridFieldInfo
    {
        public string FieldName { get; set; }
        public string Alias { get; set; }
        public string Unit { get; set; }
        public string ArchiveType { get; set; }
    }
}