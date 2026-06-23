using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace K450DataExporter
{
    public partial class Form1 : Form
    {
        private string currentInfPath;
        private string currentMdbPath;
        private string currentMdbCopyPath;
        private InfParser infParser;
        private DateTime firstDate;
        private DateTime lastDate;
        private string firstTime;
        private string lastTime;
        private int firstCycle;
        private int lastCycle;
        private DateTime lastDateMinus7Days;

        // Перечисление режимов
        private enum SelectionMode { Period, Cycle }
        private enum ArchiveType { RealTime, Historical }

        // Класс параметров запроса
        private class QueryParameters
        {
            public bool IsOdr { get; set; }
            public SelectionMode Mode { get; set; }
            public ArchiveType Archive { get; set; }
            public DateTime FromDate { get; set; }
            public DateTime ToDate { get; set; }
            public string FromTime { get; set; }
            public string ToTime { get; set; }
            public int FromCycle { get; set; }
            public int ToCycle { get; set; }
            public List<GridFieldInfo> GridFields { get; set; }
            public string TableName { get; set; }
        }

        #region Конструктор и инициализация

        public Form1(string infPath)
        {
            InitializeComponent();
            currentInfPath = infPath;

            infParser = new InfParser(currentInfPath);
            string baseDir = Path.GetDirectoryName(currentInfPath);
            currentMdbPath = infParser.GetAvailableDatabasePath(baseDir);

            if (!string.IsNullOrEmpty(infParser.NomeFileCopia))
            {
                currentMdbCopyPath = Path.Combine(baseDir, infParser.NomeFileCopia);
            }

            this.Text = $"Data Report Export - {Path.GetFileName(currentInfPath)}";

            string dbStatus = File.Exists(currentMdbPath) ? "OK" : "NOT FOUND";
            this.Text += $" | DB: {Path.GetFileName(currentMdbPath)} [{dbStatus}]";

            if (!File.Exists(currentMdbPath) && !string.IsNullOrEmpty(currentMdbCopyPath) && File.Exists(currentMdbCopyPath))
            {
                this.Text += " (using copy)";
            }

            AttachPlaceholders();
            AttachEventHandlers();
            InitializeUiState();
            LoadBoundaryValues();
        }

        private void AttachEventHandlers()
        {
            btnShowODR.Click += BtnShowODR_Click;
            btnShowPDR.Click += BtnShowPDR_Click;
            btnExportODR.Click += BtnExportODR_Click;
            btnExportPDR.Click += BtnExportPDR_Click;
            btnCancel.Click += BtnCancel_Click;

            radioSelPeriod.CheckedChanged += RadioSel_CheckedChanged;
            radioSelCycle.CheckedChanged += RadioSel_CheckedChanged;

            btnFirstDateODR.Click += (s, e) => SetFirstLastDate(textFromDateODR, true);
            btnLastDateODR.Click += (s, e) => SetFirstLastDate(textToDateODR, false);
            btnFirstTimeODR.Click += (s, e) => SetFirstLastTime(textFromTimeODR, true);
            btnLastTimeODR.Click += (s, e) => SetFirstLastTime(textToTimeODR, false);
            btnFirstCycleODR.Click += (s, e) => SetFirstLastCycle(textFromCycODR, true);
            btnLastCycleODR.Click += (s, e) => SetFirstLastCycle(textToCycODR, false);
            btn7dDateODR.Click += (s, e) => SetDateMinus7DaysToField(textFromDateODR);
            btn7dDatePDR.Click += (s, e) => SetDateMinus7DaysToField(textFromDatePDR);

            btnFirstDatePDR.Click += (s, e) => SetFirstLastDate(textFromDatePDR, true);
            btnLastDatePDR.Click += (s, e) => SetFirstLastDate(textToDatePDR, false);
            btnFirstTimePDR.Click += (s, e) => SetFirstLastTime(textFromTimePDR, true);
            btnLastTimePDR.Click += (s, e) => SetFirstLastTime(textToTimePDR, false);
            btnDiagnose.Click += BtnDiagnose_Click;
            btnPrintODR.Click += btnPrintODR_Click;
            btnPrintPDR.Click += btnPrintPDR_Click;
        }

        private void AttachPlaceholders()
        {
            AttachPlaceholderToTextBox(textFromDateODR, "YYYY-MM-DD");
            AttachPlaceholderToTextBox(textToDateODR, "YYYY-MM-DD");
            AttachPlaceholderToTextBox(textFromTimeODR, "HH:MM");
            AttachPlaceholderToTextBox(textToTimeODR, "HH:MM");
            AttachPlaceholderToTextBox(textFromCycODR, "Cycle N");
            AttachPlaceholderToTextBox(textToCycODR, "Cycle N");
            AttachPlaceholderToTextBox(textFromDatePDR, "YYYY-MM-DD");
            AttachPlaceholderToTextBox(textToDatePDR, "YYYY-MM-DD");
            AttachPlaceholderToTextBox(textFromTimePDR, "HH:MM");
            AttachPlaceholderToTextBox(textToTimePDR, "HH:MM");
        }

        private void AttachPlaceholderToTextBox(TextBox textBox, string placeholder)
        {
            textBox.Tag = placeholder;
            textBox.Text = placeholder;
            textBox.ForeColor = Color.Gray;
            textBox.Enter += (s, e) => RemovePlaceholder(textBox, placeholder);
            textBox.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                    SetPlaceholder(textBox, placeholder);
            };
        }

        private void SetPlaceholder(TextBox textBox, string placeholderText)
        {
            if (string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = placeholderText;
                textBox.ForeColor = Color.Gray;
            }
        }

        private void RemovePlaceholder(TextBox textBox, string placeholderText)
        {
            if (textBox.Text == placeholderText)
            {
                textBox.Text = "";
                textBox.ForeColor = Color.Black;
            }
        }

        private void InitializeUiState()
        {
            radioSelPeriod.Checked = true;
            radioArTRT.Checked = true;
            UpdateFieldsState();
        }

        private void UpdateFieldsState()
        {
            bool isPeriodMode = radioSelPeriod.Checked;

            textFromDateODR.Enabled = isPeriodMode;
            textToDateODR.Enabled = isPeriodMode;
            textFromTimeODR.Enabled = isPeriodMode;
            textToTimeODR.Enabled = isPeriodMode;
            textFromCycODR.Enabled = !isPeriodMode;
            textToCycODR.Enabled = !isPeriodMode;

            btnFirstDateODR.Enabled = isPeriodMode;
            btnLastDateODR.Enabled = isPeriodMode;
            btnFirstTimeODR.Enabled = isPeriodMode;
            btnLastTimeODR.Enabled = isPeriodMode;
            btnFirstCycleODR.Enabled = !isPeriodMode;
            btnLastCycleODR.Enabled = !isPeriodMode;
            btn7dDateODR.Enabled = isPeriodMode;
            btn7dDatePDR.Enabled = true;
        }

        #endregion

        #region Работа с базой данных (общие методы)

        private string GetAvailableDatabasePath()
        {
            if (File.Exists(currentMdbPath))
                return currentMdbPath;

            if (!string.IsNullOrEmpty(currentMdbCopyPath) && File.Exists(currentMdbCopyPath))
            {
                MessageBox.Show($"Основной файл базы не найден:\n{currentMdbPath}\n\nИспользуется резервная копия:\n{currentMdbCopyPath}",
                                "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return currentMdbCopyPath;
            }

            return currentMdbPath;
        }

        private string GetConnectionString()
        {
            string dbPath = GetAvailableDatabasePath();
            if (!File.Exists(dbPath)) return null;

            string connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
            try
            {
                using (var testConn = new OleDbConnection(connString))
                {
                    testConn.Open();
                }
                return connString;
            }
            catch
            {
                connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};";
                using (var testConn = new OleDbConnection(connString))
                {
                    testConn.Open();
                }
                return connString;
            }
        }

        private DataTable ExecuteQuery(string sqlQuery)
        {
            string connString = GetConnectionString();
            if (string.IsNullOrEmpty(connString)) return null;

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    return dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка запроса:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        #endregion

        #region Валидация данных

        private bool IsValidDate(TextBox textBox, out DateTime result)
        {
            string placeholder = textBox.Tag?.ToString();
            string text = textBox.Text;

            if (string.IsNullOrWhiteSpace(text) || text == placeholder)
            {
                result = DateTime.MinValue;
                return false;
            }

            return DateTime.TryParseExact(text, "yyyy-MM-dd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out result);
        }

        private bool IsValidTime(TextBox textBox, out TimeSpan result)
        {
            string placeholder = textBox.Tag?.ToString();
            string text = textBox.Text;
            result = TimeSpan.Zero;

            if (string.IsNullOrWhiteSpace(text) || text == placeholder)
                return false;

            if (DateTime.TryParseExact(text, "HH:mm",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out DateTime dt))
            {
                result = dt.TimeOfDay;
                return true;
            }

            return false;
        }

        private bool IsValidCycle(TextBox textBox, out int result)
        {
            string placeholder = textBox.Tag?.ToString();
            string text = textBox.Text;
            result = 0;

            if (string.IsNullOrWhiteSpace(text) || text == placeholder)
                return false;

            return int.TryParse(text, out result);
        }

        private bool ValidateOdrParameters(SelectionMode mode, out QueryParameters parameters)
        {
            parameters = new QueryParameters
            {
                IsOdr = true,
                Mode = mode,
                Archive = radioArTRT.Checked ? ArchiveType.RealTime : ArchiveType.Historical,
                GridFields = infParser.GetGridProcessFields(),
                TableName = InfParser.ProcessTableName
            };

            if (mode == SelectionMode.Period)
            {
                if (!IsValidDate(textFromDateODR, out DateTime fromDate) ||
                    !IsValidDate(textToDateODR, out DateTime toDate))
                {
                    MessageBox.Show("Неверный формат даты.\nОжидаемый формат: ГГГГ-ММ-ДД",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (fromDate > toDate)
                {
                    MessageBox.Show("Начальная дата не может быть больше конечной.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                DateTime today = DateTime.Now.Date;
                if (fromDate > today || toDate > today)
                {
                    MessageBox.Show("Дата не может быть позже сегодняшнего дня.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (!IsValidTime(textFromTimeODR, out TimeSpan fromTime) ||
                    !IsValidTime(textToTimeODR, out TimeSpan toTime))
                {
                    MessageBox.Show("Неверный формат времени.\nОжидаемый формат: ЧЧ:ММ",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (fromDate == toDate && fromTime > toTime)
                {
                    MessageBox.Show("Начальное время не может быть больше конечного при одинаковых датах.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                parameters.FromDate = fromDate;
                parameters.ToDate = toDate;
                parameters.FromTime = fromTime.ToString(@"hh\:mm");
                parameters.ToTime = toTime.ToString(@"hh\:mm");
            }
            else
            {
                if (!IsValidCycle(textFromCycODR, out int fromCycle) ||
                    !IsValidCycle(textToCycODR, out int toCycle))
                {
                    MessageBox.Show("Неверный формат номера цикла.\nОжидаемый формат: число",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (fromCycle > toCycle)
                {
                    MessageBox.Show("Начальный номер цикла не может быть больше конечного.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (fromCycle <= 0)
                {
                    MessageBox.Show("Номер цикла должен быть больше 0.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                parameters.FromCycle = fromCycle;
                parameters.ToCycle = toCycle;
            }

            return true;
        }

        private bool ValidatePdrParameters(out QueryParameters parameters)
        {
            parameters = new QueryParameters
            {
                IsOdr = false,
                Mode = SelectionMode.Period,
                Archive = ArchiveType.RealTime,
                TableName = InfParser.ProcessTableName
            };

            if (!IsValidDate(textFromDatePDR, out DateTime fromDate) ||
                !IsValidDate(textToDatePDR, out DateTime toDate))
            {
                MessageBox.Show("Неверный формат даты (PDR).\nОжидаемый формат: ГГГГ-ММ-ДД",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (fromDate > toDate)
            {
                MessageBox.Show("Начальная дата не может быть больше конечной (PDR).",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            DateTime today = DateTime.Now.Date;
            if (fromDate > today || toDate > today)
            {
                MessageBox.Show("Дата не может быть позже сегодняшнего дня (PDR).",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!IsValidTime(textFromTimePDR, out TimeSpan fromTime) ||
                !IsValidTime(textToTimePDR, out TimeSpan toTime))
            {
                MessageBox.Show("Неверный формат времени (PDR).\nОжидаемый формат: ЧЧ:ММ",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (fromDate == toDate && fromTime > toTime)
            {
                MessageBox.Show("Начальное время не может быть больше конечного при одинаковых датах (PDR).",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            parameters.FromDate = fromDate;
            parameters.ToDate = toDate;
            parameters.FromTime = fromTime.ToString(@"hh\:mm");
            parameters.ToTime = toTime.ToString(@"hh\:mm");

            return true;
        }

        #endregion

        #region Построение SQL запросов и вычисление формул

        /// </summary>
        private class ParsedFormula
        {
            public List<string> SqlParts { get; set; } = new List<string>();  // SUM(PRC002), SUM(PRC003) и т.д.
            public string CSharpExpression { get; set; } = "";                 // {0} / {1} * 1000 и т.д.
            public List<string> FieldNames { get; set; } = new List<string>(); // PRC002, PRC003
        }

        /// <summary>
        /// Универсальный парсер формул
        /// </summary>
        private ParsedFormula ParseFormula(string formula)
        {
            var result = new ParsedFormula();
            string tempFormula = formula;
            int placeholderIndex = 0;

            // Находим все SUM(XXX), AVG(XXX), MIN(XXX), MAX(XXX), COUNT(XXX)
            string[] functions = { "SUM(", "AVG(", "MIN(", "MAX(", "COUNT(" };

            foreach (var func in functions)
            {
                int startIndex = 0;
                while (startIndex < tempFormula.Length && (startIndex = tempFormula.IndexOf(func, startIndex)) != -1)
                {
                    int endIndex = tempFormula.IndexOf(")", startIndex);
                    if (endIndex == -1) break;

                    // Проверяем, что между func и закрывающей скобкой есть что-то
                    if (endIndex <= startIndex + func.Length)
                    {
                        // Пустые скобки SUM() - пропускаем
                        startIndex = endIndex + 1;
                        continue;
                    }

                    // Проверяем, что индексы корректны
                    if (startIndex + func.Length >= tempFormula.Length)
                    {
                        startIndex = endIndex + 1;
                        continue;
                    }

                    string fullExpression = tempFormula.Substring(startIndex, endIndex - startIndex + 1);

                    // Проверяем длину fullExpression
                    if (fullExpression.Length <= func.Length)
                    {
                        startIndex = endIndex + 1;
                        continue;
                    }

                    string fieldName = fullExpression.Substring(func.Length, fullExpression.Length - func.Length - 1);

                    result.SqlParts.Add(fullExpression);
                    result.FieldNames.Add(fieldName);

                    // Заменяем на плейсхолдер
                    tempFormula = tempFormula.Replace(fullExpression, $"{{{placeholderIndex}}}");
                    placeholderIndex++;

                    // После замены строка стала короче, продолжаем с текущей позиции
                    startIndex = startIndex + 1;
                }
            }

            result.CSharpExpression = tempFormula;
            return result;
        }

        /// <summary>
        /// Собирает все SQL части из списка формул
        /// </summary>
        private HashSet<string> CollectAllSqlParts(List<ReportField> reportFields)
        {
            var allSqlParts = new HashSet<string>();

            foreach (var reportField in reportFields)
            {
                var parsed = ParseFormula(reportField.Formula);
                foreach (var sqlPart in parsed.SqlParts)
                {
                    allSqlParts.Add(sqlPart);
                }
            }

            return allSqlParts;
        }

        /// <summary>
        /// Строит SQL запрос на основе агрегатных функций
        /// </summary>
        private string BuildDynamicSqlQuery(string whereClause, List<ReportField> reportFields)
        {
            var allSqlParts = CollectAllSqlParts(reportFields);

            if (allSqlParts.Count == 0)
            {
                return "";
            }

            var selectParts = new List<string>();
            int idx = 0;

            foreach (var sqlPart in allSqlParts)
            {
                // Извлекаем имя поля из SUM(PRC002)
                string fieldName = sqlPart.Substring(sqlPart.IndexOf("(") + 1, sqlPart.LastIndexOf(")") - sqlPart.IndexOf("(") - 1);
                string cleanFieldName = fieldName.Replace("*", "_");

                // Обрабатываем NULL через IIF
                if (fieldName.Contains("*"))
                {
                    string[] parts = fieldName.Split('*');
                    if (parts.Length == 2)
                    {
                        selectParts.Add($"SUM(IIF({parts[0]} IS NULL, 0, {parts[0]}) * IIF({parts[1]} IS NULL, 0, {parts[1]})) AS [{sqlPart}]");
                    }
                    else
                    {
                        selectParts.Add($"SUM(IIF({fieldName} IS NULL, 0, {fieldName})) AS [{sqlPart}]");
                    }
                }
                else
                {
                    selectParts.Add($"SUM(IIF({fieldName} IS NULL, 0, {fieldName})) AS [{sqlPart}]");
                }
                idx++;
            }

            string selectFields = string.Join(", ", selectParts);

            return $@"
        SELECT TOP 50000 
            {selectFields}
        FROM [Process_DataE]
        WHERE {whereClause}";
        }

        /// <summary>
        /// Вычисляет формулу, подставляя значения из DataRow
        /// </summary>
        private string CalculateFormulaFromData(string formula, DataRow dataRow)
        {
            DataTable dt = new DataTable();
            if (string.IsNullOrEmpty(formula)) return "0";

            try
            {
                var parsed = ParseFormula(formula);

                if (parsed.SqlParts.Count == 0)
                {
                    var result = dt.Compute(formula, "");
                    return Convert.ToDouble(result).ToString("0.###");
                }

                // Подставляем значения из DataRow
                string expression = parsed.CSharpExpression;
                for (int i = 0; i < parsed.SqlParts.Count; i++)
                {
                    string sqlPart = parsed.SqlParts[i];
                    string value = "0";

                    if (dataRow.Table.Columns.Contains(sqlPart) && dataRow[sqlPart] != DBNull.Value)
                    {
                        value = dataRow[sqlPart].ToString().Replace(",", ".");
                    }

                    expression = expression.Replace($"{{{i}}}", value);
                }

                // Вычисляем выражение

                var computedValue = dt.Compute(expression, "");
                return Convert.ToDouble(computedValue).ToString("0.###");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка: {ex.Message}");
                return "0";
            }
        }



        /// <summary>
        /// Раскрывает PRDxxx в формулу из PRODUCT_TABLE
        /// </summary>
        private string ExpandPrdToPrc(string formula)
        {
            string result = formula;
            int maxIterations = 10;

            for (int iter = 0; iter < maxIterations && result.Contains("PRD"); iter++)
            {
                foreach (var productField in infParser.ProductFields)
                {
                    string prdName = productField.Name;
                    if (result.Contains(prdName))
                    {
                        string prdFormula = productField.Formula; // SUM(PRC002)/1000
                                                                  // Добавляем скобки вокруг подставляемой формулы
                        result = result.Replace(prdName, $"({prdFormula})");
                    }
                }
            }

            return result;
        }

        #endregion

        #region Отображение данных

        private void DisplayOdrRealtimeTable(DataGridView grid, DataTable data, QueryParameters parameters, List<GridFieldInfo> gridFields)
        {
            labelOdrTitle.Text = GetOdrQueryInfoText(parameters);
            labelOdrTitle.Visible = true;
            grid.DataSource = data;

            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            if (grid.Columns.Contains("Data"))
                grid.Columns["Data"].HeaderText = "Data";
            if (grid.Columns.Contains("Ora"))
            {
                grid.Columns["Ora"].HeaderText = "Ora";
                grid.Columns["Ora"].DefaultCellStyle.Format = "HH:mm";
            }
            if (grid.Columns.Contains("Cy"))
                grid.Columns["Cy"].HeaderText = "Cycle";

            int fieldIndex = 0;
            foreach (DataGridViewColumn col in grid.Columns)
            {
                if (col.Name == "Data" || col.Name == "Ora" || col.Name == "Cy") continue;
                if (fieldIndex < gridFields.Count)
                {
                    var field = gridFields[fieldIndex];
                    string header = field.Alias;
                    if (!string.IsNullOrEmpty(field.Unit) && field.Unit != "#" && field.Unit != "null")
                    {
                        header = $"{field.Alias} ({field.Unit})";
                    }
                    col.HeaderText = header;
                    fieldIndex++;
                }
            }
        }

        #endregion

        #region Основные операции

        private void ShowOdrData()
        {
            if (!File.Exists(GetAvailableDatabasePath()))
            {
                MessageBox.Show("Файл базы данных не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SelectionMode mode = radioSelPeriod.Checked ? SelectionMode.Period : SelectionMode.Cycle;
            ArchiveType archive = radioArTRT.Checked ? ArchiveType.RealTime : ArchiveType.Historical;

            if (!ValidateOdrParameters(mode, out QueryParameters parameters))
                return;

            string whereClause = (parameters.Mode == SelectionMode.Period)
                ? $"[Data] BETWEEN #{parameters.FromDate:MM/dd/yyyy}# AND #{parameters.ToDate:MM/dd/yyyy}# AND [Ora] BETWEEN #{parameters.FromTime}# AND #{parameters.ToTime}#"
                : $"[Cy] BETWEEN {parameters.FromCycle} AND {parameters.ToCycle}";

            if (archive == ArchiveType.RealTime)
            {
                var dbColumns = new List<string> { "[Data]", "[Ora]", "[Cy]" };
                foreach (var field in parameters.GridFields)
                {
                    dbColumns.Add($"[{field.FieldName}]");
                }
                string selectFields = string.Join(", ", dbColumns);
                string sqlQuery = $@"
            SELECT {selectFields}
            FROM [Process_DataE]
            WHERE {whereClause}
            ORDER BY [Data] DESC, [Ora] DESC";

                DataTable data = ExecuteQuery(sqlQuery);
                if (data != null)
                {
                    DisplayOdrRealtimeTable(dataTableODR, data, parameters, parameters.GridFields);

                    // --- УПРАВЛЕНИЕ КНОПКАМИ ---
                    bool hasRows = data.Rows.Count > 0;
                    btnExportODR.Enabled = hasRows;
                    btnPrintODR.Enabled = hasRows; // Включаем/выключаем печать ODR

                    MessageBox.Show($"Загружено записей: {data.Rows.Count}", "Готово",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    btnPrintODR.Enabled = false; // На всякий случай гасим, если запрос вернул null
                }
            }
            else
            {
                string dynamicQuery = BuildDynamicSqlQuery(whereClause, infParser.ProcessReportFields);

                if (string.IsNullOrEmpty(dynamicQuery))
                {
                    MessageBox.Show("Не удалось построить запрос для формул.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataTable historicalData = ExecuteQuery(dynamicQuery);

                if (historicalData != null && historicalData.Rows.Count > 0)
                {
                    DataRow row = historicalData.Rows[0];
                    DataTable reportTable = new DataTable();
                    reportTable.Columns.Add("Параметр", typeof(string));
                    reportTable.Columns.Add("Значение", typeof(string));
                    reportTable.Columns.Add("Ед. изм.", typeof(string));

                    foreach (var reportField in infParser.ProcessReportFields)
                    {
                        string value = CalculateFormulaFromData(reportField.Formula, row);
                        reportTable.Rows.Add(reportField.Name, value, reportField.Unit);
                    }

                    dataTableODR.DataSource = reportTable;

                    // --- УПРАВЛЕНИЕ КНОПКАМИ ---
                    btnExportODR.Enabled = true;
                    btnPrintODR.Enabled = true; // Включаем печать для исторического отчета ODR

                    dataTableODR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    if (dataTableODR.Columns.Count >= 3)
                    {
                        dataTableODR.Columns[0].Width = 250;
                        dataTableODR.Columns[1].Width = 120;
                        dataTableODR.Columns[2].Width = 100;
                    }
                }
                else
                {
                    // Если исторических данных нет — гасим обе кнопки
                    btnExportODR.Enabled = false;
                    btnPrintODR.Enabled = false;
                }
            }
        }

        private void ShowPdrData()
        {
            if (!File.Exists(GetAvailableDatabasePath()))
            {
                MessageBox.Show("Файл базы данных не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!ValidatePdrParameters(out QueryParameters parameters))
                return;

            string whereClause = $"[Data] BETWEEN #{parameters.FromDate:MM/dd/yyyy}# AND #{parameters.ToDate:MM/dd/yyyy}# " +
                                 $"AND [Ora] BETWEEN #{parameters.FromTime}# AND #{parameters.ToTime}#";

            var simplifiedFormulas = new List<string>();
            foreach (var reportField in infParser.ProductReportFields)
            {
                string expanded = ExpandPrdToPrc(reportField.Formula);
                string simplified = SimplifyFormula(expanded);

                simplified = simplified.Replace("*1", "");
                simplified = simplified.Replace("1*", "");

                simplifiedFormulas.Add(simplified);
            }

            var requiredPrcFields = new HashSet<string>();
            foreach (string formula in simplifiedFormulas)
            {
                var fields = ExtractPrcFields(formula);
                foreach (var field in fields)
                {
                    requiredPrcFields.Add(field);
                }
            }

            var selectParts = new List<string>();
            foreach (string field in requiredPrcFields)
            {
                if (field.Contains("*"))
                {
                    string[] parts = field.Split('*');
                    selectParts.Add($"SUM(IIF({parts[0]} IS NULL, 0, {parts[0]}) * IIF({parts[1]} IS NULL, 0, {parts[1]})) AS SUM_{field.Replace("*", "_")}");
                }
                else
                {
                    selectParts.Add($"SUM(IIF({field} IS NULL, 0, {field})) AS SUM_{field}");
                }
            }

            string selectFields = string.Join(", ", selectParts);
            string sqlQuery = $@"
        SELECT TOP 50000 
            {selectFields}
        FROM [Process_DataE]
        WHERE {whereClause}";

            DataTable data = ExecuteQuery(sqlQuery);

            if (data != null && data.Rows.Count > 0)
            {
                DataRow row = data.Rows[0];
                DataTable reportTable = new DataTable();
                reportTable.Columns.Add("Параметр", typeof(string));
                reportTable.Columns.Add("Значение", typeof(string));
                reportTable.Columns.Add("Ед. изм.", typeof(string));

                for (int i = 0; i < infParser.ProductReportFields.Count; i++)
                {
                    var reportField = infParser.ProductReportFields[i];
                    string simplifiedFormula = simplifiedFormulas[i];
                    double value = EvaluatePdrFormula(simplifiedFormula, row);
                    reportTable.Rows.Add(reportField.Name, value.ToString("0.###"), reportField.Unit);
                }

                dataTablePDR.DataSource = reportTable;

                // --- УПРАВЛЕНИЕ КНОПКАМИ ---
                btnExportPDR.Enabled = true;
                btnPrintPDR.Enabled = true; // Включаем печать для PDR отчета

                dataTablePDR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                if (dataTablePDR.Columns.Count >= 3)
                {
                    dataTablePDR.Columns[0].Width = 250;
                    dataTablePDR.Columns[1].Width = 120;
                    dataTablePDR.Columns[2].Width = 100;
                }
            }
            else
            {
                // Если данных по PDR нет — блокируем кнопки
                btnExportPDR.Enabled = false;
                btnPrintPDR.Enabled = false;
            }
        }

        private double EvaluatePdrFormula(string formula, DataRow row)
        {
            string expression = formula;

            // 1. Заменяем PRCxxx на значения из базы
            var matches = System.Text.RegularExpressions.Regex.Matches(expression, @"PRC[0-9]+(?:\*PRC[0-9]+)?");
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                string field = match.Value;
                string columnName = $"SUM_{field.Replace("*", "_")}";

                string value = "0";
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    value = row[columnName].ToString().Replace(",", ".");
                }
                expression = System.Text.RegularExpressions.Regex.Replace(expression, @"/0\b", "/1");
                expression = expression.Replace(field, value);
            }

            // 2. Удаляем все оставшиеся SUM( ... ) — они мешают Compute
            expression = System.Text.RegularExpressions.Regex.Replace(expression, @"SUM\(([^)]+)\)", "$1");

            // 3. Вычисляем
            try
            {
                DataTable dt = new DataTable();
                var computedValue = dt.Compute(expression, "");
                return Convert.ToDouble(computedValue);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка: {expression}");
                System.Diagnostics.Debug.WriteLine(ex.Message);
                return 0;
            }
        }
        private HashSet<string> ExtractPrcFields(string formula)
        {
            var fields = new HashSet<string>();
            var matches = System.Text.RegularExpressions.Regex.Matches(formula, @"PRC[0-9]+(?:\*PRC[0-9]+)?");

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                fields.Add(match.Value);
            }

            return fields;
        }
        private string SimplifyFormula(string formula)
        {
            string result = formula;

            // Убираем ВСЕ SUM( ... ) и оставляем содержимое
            // SUM(PRC002/1000) -> PRC002/1000
            // SUM(PRC003*1) -> PRC003*1
            // SUM(PRC003*1)/SUM(PRC002/1000) -> PRC003*1/PRC002/1000

            while (result.Contains("SUM("))
            {
                result = System.Text.RegularExpressions.Regex.Replace(result, @"SUM\(([^)]+)\)", "$1");
            }

            return result;
        }
        private void ExportOdrData()
        {
            if (!File.Exists(GetAvailableDatabasePath()))
            {
                MessageBox.Show("Файл базы данных не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SelectionMode mode = radioSelPeriod.Checked ? SelectionMode.Period : SelectionMode.Cycle;
            ArchiveType archive = radioArTRT.Checked ? ArchiveType.RealTime : ArchiveType.Historical;

            if (!ValidateOdrParameters(mode, out QueryParameters parameters))
                return;

            // Определяем whereClause один раз для обоих случаев
            string whereClause = (parameters.Mode == SelectionMode.Period)
                ? $"[Data] BETWEEN #{parameters.FromDate:MM/dd/yyyy}# AND #{parameters.ToDate:MM/dd/yyyy}# AND [Ora] BETWEEN #{parameters.FromTime}# AND #{parameters.ToTime}#"
                : $"[Cy] BETWEEN {parameters.FromCycle} AND {parameters.ToCycle}";

            if (archive == ArchiveType.RealTime)
            {
                // --- РЕЖИМ REALTIME (Широкая таблица, много колонок PRCxxx) ---
                var dbColumns = new List<string> { "[Data]", "[Ora]", "[Cy]" };
                foreach (var field in parameters.GridFields)
                {
                    dbColumns.Add($"[{field.FieldName}]");
                }
                string selectFields = string.Join(", ", dbColumns);
                string sqlQuery = $@"
            SELECT {selectFields}
            FROM [Process_DataE]
            WHERE {whereClause}
            ORDER BY [Data] DESC, [Ora] DESC";

                DataTable data = ExecuteQuery(sqlQuery);
                if (data != null && data.Rows.Count > 0)
                {
                    // Здесь работает модифицированный SaveDataTableToCsv (который сопоставляет заголовки колонок)
                    SaveDataTableToCsv(data, parameters, "ODR");
                    btnExportODR.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Нет данных для экспорта по заданным параметрам.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                // --- РЕЖИМ HISTORICAL (Вертикальный отчет: Параметр, Значение, Ед.изм) ---
                string dynamicQuery = BuildDynamicSqlQuery(whereClause, infParser.ProcessReportFields);

                if (string.IsNullOrEmpty(dynamicQuery))
                {
                    MessageBox.Show("Не удалось построить запрос для формул.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataTable historicalData = ExecuteQuery(dynamicQuery);

                if (historicalData != null && historicalData.Rows.Count > 0 && historicalData.Rows[0][0] != DBNull.Value)
                {
                    DataRow row = historicalData.Rows[0];
                    DataTable reportTable = new DataTable();
                    reportTable.Columns.Add("Параметр", typeof(string));
                    reportTable.Columns.Add("Значение", typeof(string));
                    reportTable.Columns.Add("Ед. изм.", typeof(string));

                    foreach (var reportField in infParser.ProcessReportFields)
                    {
                        string value = CalculateFormulaFromData(reportField.Formula, row);

                        // Формируем понятное имя для первой ячейки строки.
                        // Например, вместо "PRC029" запишет "PRC029 (Температура бака)"
                        string displayName = reportField.Name;
                        if (!string.IsNullOrEmpty(reportField.Alias) && reportField.Alias != reportField.Name)
                        {
                            displayName = $"{reportField.Name} ({reportField.Alias})";

                            // Если тебе в CSV нужен ТОЛЬКО комментарий без PRCxxx, закомментируй строку выше и раскомментируй строку ниже:
                            // displayName = reportField.Alias;
                        }

                        reportTable.Rows.Add(displayName, value, reportField.Unit);
                    }

                    // Вызываем экспорт трехколоночной таблицы, где внутри ячеек уже лежат комментарии
                    SaveReportTableToCsv(reportTable, parameters, "ODR");
                }
                else
                {
                    MessageBox.Show("Нет данных для экспорта по заданным параметрам.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void ExportPdrData()
        {
            if (!File.Exists(GetAvailableDatabasePath()))
            {
                MessageBox.Show("Файл базы данных не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!ValidatePdrParameters(out QueryParameters parameters))
                return;

            string whereClause = $"[Data] BETWEEN #{parameters.FromDate:MM/dd/yyyy}# AND #{parameters.ToDate:MM/dd/yyyy}# " +
                                 $"AND [Ora] BETWEEN #{parameters.FromTime}# AND #{parameters.ToTime}#";

            // ШАГ 1: Упрощаем формулы (как в ShowPdrData)
            var simplifiedFormulas = new List<string>();
            foreach (var reportField in infParser.ProductReportFields)
            {
                string expanded = ExpandPrdToPrc(reportField.Formula);
                string simplified = SimplifyFormula(expanded);
                simplified = simplified.Replace("*1", "").Replace("1*", "");
                simplifiedFormulas.Add(simplified);
            }

            // ШАГ 2: Собираем все уникальные PRC поля из упрощённых формул
            var requiredPrcFields = new HashSet<string>();
            foreach (string formula in simplifiedFormulas)
            {
                var fields = ExtractPrcFields(formula);
                foreach (var field in fields)
                {
                    requiredPrcFields.Add(field);
                }
            }

            // ШАГ 3: Строим SQL запрос
            var selectParts = new List<string>();
            foreach (string field in requiredPrcFields)
            {
                if (field.Contains("*"))
                {
                    string[] parts = field.Split('*');
                    selectParts.Add($"SUM(IIF({parts[0]} IS NULL, 0, {parts[0]}) * IIF({parts[1]} IS NULL, 0, {parts[1]})) AS SUM_{field.Replace("*", "_")}");
                }
                else
                {
                    selectParts.Add($"SUM(IIF({field} IS NULL, 0, {field})) AS SUM_{field}");
                }
            }

            string selectFields = string.Join(", ", selectParts);
            string sqlQuery = $@"
        SELECT TOP 50000 
            {selectFields}
        FROM [Process_DataE]
        WHERE {whereClause}";

            DataTable data = ExecuteQuery(sqlQuery);

            if (data != null && data.Rows.Count > 0 && data.Rows[0][0] != DBNull.Value)
            {
                DataRow row = data.Rows[0];
                DataTable reportTable = new DataTable();
                reportTable.Columns.Add("Параметр", typeof(string));
                reportTable.Columns.Add("Значение", typeof(string));
                reportTable.Columns.Add("Ед. изм.", typeof(string));

                // ШАГ 4: Вычисляем значения (как в ShowPdrData)
                for (int i = 0; i < infParser.ProductReportFields.Count; i++)
                {
                    var reportField = infParser.ProductReportFields[i];
                    string simplifiedFormula = simplifiedFormulas[i];
                    double value = EvaluatePdrFormula(simplifiedFormula, row);
                    reportTable.Rows.Add(reportField.Name, value.ToString("0.###"), reportField.Unit);
                }

                SaveReportTableToCsv(reportTable, parameters, "PDR");
            }
            else
            {
                MessageBox.Show("Нет данных для экспорта по заданным параметрам.",
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SaveDataTableToCsv(DataTable data, QueryParameters parameters, string reportType)
        {
            string exportFolder = Path.Combine(Path.GetDirectoryName(currentInfPath), "Reports");
            if (!Directory.Exists(exportFolder))
                Directory.CreateDirectory(exportFolder);

            string fileName = GenerateFileName(parameters, reportType);
            string filePath = Path.Combine(exportFolder, fileName);

            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    string queryInfo = (reportType == "ODR") ? GetOdrQueryInfoText(parameters) : GetPdrQueryInfoText(parameters);
                    sw.WriteLine($"# {queryInfo}");
                    sw.WriteLine("#");

                    // =================================================================
                    // 1. БЛОК ЗАПИСИ ЗАГОЛОВКОВ С КОММЕНТАРИЯМИ
                    // =================================================================
                    for (int i = 0; i < data.Columns.Count; i++)
                    {
                        string columnName = data.Columns[i].ColumnName;
                        string headerText = columnName;

                        // Исключаем системные колонки времени и циклов, их переводить не нужно
                        if (columnName != "Data" && columnName != "Ora" && columnName != "Cy")
                        {
                            // Ищем описание поля в параметрах GRID
                            if (parameters.GridFields != null)
                            {
                                var fieldInfo = parameters.GridFields.FirstOrDefault(f => f.FieldName.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                                if (fieldInfo != null)
                                {
                                    // Формируем красивое описание. Например: "PRC001 (Cycle time)"
                                    string description = !string.IsNullOrEmpty(fieldInfo.Alias) ? fieldInfo.Alias : "Без описания";

                                    // Если есть адекватная единица измерения, добавляем и её
                                    if (!string.IsNullOrEmpty(fieldInfo.Unit) && fieldInfo.Unit != "#" && fieldInfo.Unit != "null")
                                    {
                                        headerText = $"{columnName} ({description}, {fieldInfo.Unit})";
                                    }
                                    else
                                    {
                                        headerText = $"{columnName} ({description})";
                                    }
                                }
                            }
                        }
                        else if (columnName == "Cy")
                        {
                            headerText = "Cy (Номер цикла)"; // Небольшой бонус для удобства пользователя
                        }

                        sw.Write(headerText);
                        if (i < data.Columns.Count - 1) sw.Write(";");
                    }
                    sw.WriteLine();

                    // =================================================================
                    // 2. ДОБАВЛЕННЫЙ БЛОК: ЗАПИСЬ СТРОК С ДАННЫМИ
                    // =================================================================
                    foreach (DataRow row in data.Rows)
                    {
                        for (int i = 0; i < data.Columns.Count; i++)
                        {
                            object cellValue = row[i];
                            string cellText = "";

                            if (cellValue != DBNull.Value && cellValue != null)
                            {
                                // Если это дата, приводим к понятному виду без системных "хвостов"
                                if (cellValue is DateTime dt)
                                {
                                    // Если колонка называется Ora, берем только время, иначе - только дату
                                    if (data.Columns[i].ColumnName.Equals("Ora", StringComparison.OrdinalIgnoreCase))
                                        cellText = dt.ToString("HH:mm:ss");
                                    else
                                        cellText = dt.ToString("yyyy-MM-dd");
                                }
                                else
                                {
                                    cellText = cellValue.ToString();
                                }
                            }

                            sw.Write(cellText);
                            if (i < data.Columns.Count - 1) sw.Write(";");
                        }
                        sw.WriteLine();
                    }
                }

                MessageBox.Show($"Экспорт завершён!\n\nФайл: {filePath}", "Успех",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start("explorer.exe", exportFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveReportTableToCsv(DataTable reportTable, QueryParameters parameters, string reportType)
        {
            string exportFolder = Path.Combine(Path.GetDirectoryName(currentInfPath), "Reports");
            if (!Directory.Exists(exportFolder))
                Directory.CreateDirectory(exportFolder);

            string fileName = GenerateFileName(parameters, reportType);
            string filePath = Path.Combine(exportFolder, fileName);

            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    string queryInfo = (reportType == "ODR") ? GetOdrQueryInfoText(parameters) : GetPdrQueryInfoText(parameters);
                    sw.WriteLine($"# {queryInfo}");
                    sw.WriteLine("#");

                    for (int i = 0; i < reportTable.Columns.Count; i++)
                    {
                        string columnName = reportTable.Columns[i].ColumnName;
                        string headerText = columnName;

                        if (parameters.GridFields != null)
                        {
                            var fieldInfo = parameters.GridFields.FirstOrDefault(f => f.FieldName.Equals(columnName, StringComparison.OrdinalIgnoreCase));
                            if (fieldInfo != null)
                            {
                                string description = !string.IsNullOrEmpty(fieldInfo.Alias) ? fieldInfo.Alias : "Без описания";
                                headerText = $"{columnName} ({description})";
                            }
                        }

                        sw.Write(headerText);
                        if (i < reportTable.Columns.Count - 1) sw.Write(";");
                    }
                    sw.WriteLine();

                    foreach (DataRow row in reportTable.Rows)
                    {
                        for (int i = 0; i < reportTable.Columns.Count; i++)
                        {
                            sw.Write(row[i]?.ToString());
                            if (i < reportTable.Columns.Count - 1) sw.Write(";");
                        }
                        sw.WriteLine();
                    }
                }

                MessageBox.Show($"Экспорт завершён!\n\nФайл: {filePath}", "Успех",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start("explorer.exe", exportFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetOdrQueryInfoText(QueryParameters parameters)
        {
            if (parameters.Archive == ArchiveType.RealTime)
            {
                if (parameters.Mode == SelectionMode.Period)
                {
                    return $"ODR RealTime | Период: {parameters.FromDate:yyyy-MM-dd} {parameters.FromTime} - {parameters.ToDate:yyyy-MM-dd} {parameters.ToTime}";
                }
                else
                {
                    return $"ODR RealTime | Циклы: {parameters.FromCycle} - {parameters.ToCycle}";
                }
            }
            else
            {
                if (parameters.Mode == SelectionMode.Period)
                {
                    return $"ODR Historical (отчет) | Период: {parameters.FromDate:yyyy-MM-dd} {parameters.FromTime} - {parameters.ToDate:yyyy-MM-dd} {parameters.ToTime}";
                }
                else
                {
                    return $"ODR Historical (отчет) | Циклы: {parameters.FromCycle} - {parameters.ToCycle}";
                }
            }
        }

        private string GetPdrQueryInfoText(QueryParameters parameters)
        {
            return $"PDR отчет | Период: {parameters.FromDate:yyyy-MM-dd} {parameters.FromTime} - {parameters.ToDate:yyyy-MM-dd} {parameters.ToTime}";
        }

        private string GenerateFileName(QueryParameters parameters, string reportType)
        {
            string dateTimeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string result = "";

            if (reportType == "ODR")
            {
                string modeStr = (parameters.Mode == SelectionMode.Period) ? "PR" : "CY";
                string archiveStr = (parameters.Archive == ArchiveType.RealTime) ? "RT" : "HIST";

                if (parameters.Mode == SelectionMode.Period)
                {
                    result = $"ODReport_{modeStr}_{parameters.FromDate:yyyy-MM-dd}_by_{parameters.ToDate:yyyy-MM-dd}_{archiveStr}_{dateTimeStamp}.csv";
                }
                else
                {
                    result = $"ODReport_{modeStr}_{parameters.FromCycle}_by_{parameters.ToCycle}_{archiveStr}_{dateTimeStamp}.csv";
                }
            }
            else if (reportType == "PDR")
            {
                result = $"PDReport_PR_{parameters.FromDate:yyyy-MM-dd}_by_{parameters.ToDate:yyyy-MM-dd}_{dateTimeStamp}.csv";
            }

            return result;
        }

        #endregion

        #region Обработчики кнопок

        private void BtnShowODR_Click(object sender, EventArgs e)
        {
            ShowOdrData();
        }

        private void BtnShowPDR_Click(object sender, EventArgs e)
        {
            ShowPdrData();
        }

        private void BtnExportODR_Click(object sender, EventArgs e)
        {
            ExportOdrData();
        }

        private void BtnExportPDR_Click(object sender, EventArgs e)
        {
            ExportPdrData();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void RadioSel_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFieldsState();
        }

        // Исправленный обработчик клика по кнопке "Печать ODR"
        private void btnPrintODR_Click(object sender, EventArgs e)
        {
            ExecuteTablePrint(dataTableODR, "Технологический отчет ODR");
        }

        // Исправленный обработчик клика по кнопке "Печать PDR"
        private void btnPrintPDR_Click(object sender, EventArgs e)
        {
            ExecuteTablePrint(dataTablePDR, "Производственный отчет PDR");
        }

        #endregion

        #region Универсальный многостраничный движок печати (Версия с горизонтальным переносом)
        private void ExecuteTablePrint(DataGridView grid, string reportTitle)
        {
            if (grid == null || grid.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для печати.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            System.Drawing.Printing.PrintDocument printDoc = new System.Drawing.Printing.PrintDocument();
            printDoc.DefaultPageSettings.Landscape = true;
            printDoc.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(30, 30, 30, 30);

            Font fontTitle = new Font("Arial", 11, FontStyle.Bold);
            Font fontHeader = new Font("Arial", 8.5f, FontStyle.Bold);
            Font fontBody = new Font("Arial", 8.5f, FontStyle.Regular);

            using (Graphics gMeasure = grid.CreateGraphics())
            {
                Dictionary<int, int> calculatedWidths = new Dictionary<int, int>();
                var visibleColumns = grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();

                int fixedColumnsCount = Math.Min(3, visibleColumns.Count);
                var fixedColumns = visibleColumns.Take(fixedColumnsCount).ToList();
                var dynamicColumns = visibleColumns.Skip(fixedColumnsCount).ToList();

                // 1. Предварительный расчет чистой ширины ячеек
                foreach (var col in visibleColumns)
                {
                    if (col.Name.Equals("Data", StringComparison.OrdinalIgnoreCase) || col.HeaderText.StartsWith("Data", StringComparison.OrdinalIgnoreCase))
                    {
                        calculatedWidths[col.Index] = 75;
                        continue;
                    }
                    if (col.Name.Equals("Ora", StringComparison.OrdinalIgnoreCase) || col.HeaderText.StartsWith("Ora", StringComparison.OrdinalIgnoreCase))
                    {
                        calculatedWidths[col.Index] = 65;
                        continue;
                    }
                    if (col.Name.Equals("Cy", StringComparison.OrdinalIgnoreCase) || col.HeaderText.StartsWith("Cy", StringComparison.OrdinalIgnoreCase) || col.HeaderText.StartsWith("Цикл", StringComparison.OrdinalIgnoreCase))
                    {
                        calculatedWidths[col.Index] = 45;
                        continue;
                    }

                    int maxTextWidth = (int)gMeasure.MeasureString(col.HeaderText, fontHeader).Width;

                    // Если длинный текст или есть пробелы — принудительно ужимаем под перенос слов
                    if (col.HeaderText.Contains(" ") || col.HeaderText.Length > 8)
                    {
                        maxTextWidth = (int)(maxTextWidth / 2.4);
                    }

                    int sampleRows = Math.Min(grid.Rows.Count, 30);
                    for (int i = 0; i < sampleRows; i++)
                    {
                        if (grid.Rows[i].IsNewRow) continue;
                        object val = grid.Rows[i].Cells[col.Index].Value;
                        string cellText = val?.ToString() ?? "";
                        int cellTextWidth = (int)gMeasure.MeasureString(cellText, fontBody).Width;
                        if (cellTextWidth > maxTextWidth) maxTextWidth = cellTextWidth;
                    }

                    // Минимальная ширина датчика — 50 пикселей, максимальная — 130
                    calculatedWidths[col.Index] = Math.Max(50, Math.Min(maxTextWidth + 10, 130));
                }

                int fixedWidth = fixedColumns.Sum(c => calculatedWidths[c.Index]);
                int printableWidth = printDoc.DefaultPageSettings.Bounds.Height - 60; // Общая ширина А4 альбомного листа минус поля
                int availableDynamicWidth = printableWidth - fixedWidth;

                // --- ИСПРАВЛЕНО: Максимально плотная упаковка колонок перед печатью ---
                List<List<DataGridViewColumn>> columnGroups = new List<List<DataGridViewColumn>>();
                List<DataGridViewColumn> currentGroup = new List<DataGridViewColumn>();
                int currentGroupWidth = 0;

                foreach (var col in dynamicColumns)
                {
                    int colWidth = calculatedWidths[col.Index];

                    // Теперь мы точно знаем ширину и не переносим колонки раньше времени
                    if (currentGroupWidth + colWidth > availableDynamicWidth && currentGroup.Count > 0)
                    {
                        columnGroups.Add(currentGroup);
                        currentGroup = new List<DataGridViewColumn>();
                        currentGroupWidth = 0;
                    }
                    currentGroup.Add(col);
                    currentGroupWidth += colWidth;
                }
                if (currentGroup.Count > 0) columnGroups.Add(currentGroup);
                if (columnGroups.Count == 0) columnGroups.Add(new List<DataGridViewColumn>());

                int headerHeight = 45;
                int rowHeight = 18;
                int titleHeight = 25;
                int spacing = 15;

                int actualRowsCount = grid.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
                int singleBlockDataHeight = actualRowsCount * rowHeight;
                int totalSingleBlockHeight = headerHeight + singleBlockDataHeight + spacing;

                int printableHeight = printDoc.DefaultPageSettings.Bounds.Width - 60;
                int halfPageHeight = printableHeight / 2;

                int currentGroupIndex = 0;
                int currentRowIndex = 0;

                printDoc.PrintPage += (sender, e) =>
                {
                    int x = e.MarginBounds.Left;
                    int y = e.MarginBounds.Top;

                    bool isFirstBlockOnPage = true;

                    while (currentGroupIndex < columnGroups.Count)
                    {
                        var currentDynamicGroup = columnGroups[currentGroupIndex];

                        List<DataGridViewColumn> colsToPrint = new List<DataGridViewColumn>();
                        colsToPrint.AddRange(fixedColumns);
                        colsToPrint.AddRange(currentDynamicGroup);

                        if (isFirstBlockOnPage && currentRowIndex == 0)
                        {
                            e.Graphics.DrawString($"{reportTitle} от {DateTime.Now:dd.MM.yyyy HH:mm}", fontTitle, Brushes.Black, x, y);
                            y += titleHeight;
                            isFirstBlockOnPage = false;
                        }

                        StringFormat headerFormat = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center,
                            FormatFlags = StringFormatFlags.LineLimit
                        };

                        // --- ИСПРАВЛЕНО: Равномерное распределение лишнего пространства ---
                        int totalCalculatedWidth = colsToPrint.Sum(c => calculatedWidths[c.Index]);
                        int remainingEmptySpace = e.MarginBounds.Width - totalCalculatedWidth;

                        // Считаем добавку к каждому динамическому датчику, чтобы размазать остаток пустых пикселей поровну
                        int extraWidthPerSensor = 0;
                        if (remainingEmptySpace > 0 && currentDynamicGroup.Count > 0)
                        {
                            extraWidthPerSensor = remainingEmptySpace / currentDynamicGroup.Count;
                        }

                        // Вспомогательная функция расчета финальной ширины на бумаге
                        Func<DataGridViewColumn, int> getCellWidth = (col) =>
                        {
                            int baseWidth = calculatedWidths[col.Index];

                            // Если колонка фиксированная — возвращаем её чистый размер без изменений!
                            if (fixedColumns.Contains(col)) return baseWidth;

                            // Если динамический датчик — добавляем ему его равную долю от остатка
                            return baseWidth + extraWidthPerSensor;
                        };

                        // 1. Отрисовка шапки
                        foreach (var col in colsToPrint)
                        {
                            int cellWidth = getCellWidth(col);
                            e.Graphics.DrawRectangle(Pens.Black, x, y, cellWidth, headerHeight);
                            e.Graphics.DrawString(col.HeaderText, fontHeader, Brushes.Black,
                                new RectangleF(x + 2, y + 2, cellWidth - 4, headerHeight - 4), headerFormat);
                            x += cellWidth;
                        }

                        y += headerHeight;
                        x = e.MarginBounds.Left;

                        // 2. Отрисовка строк
                        int tempRowIndex = currentRowIndex;
                        while (tempRowIndex < grid.Rows.Count)
                        {
                            DataGridViewRow row = grid.Rows[tempRowIndex];
                            if (row.IsNewRow)
                            {
                                tempRowIndex++;
                                continue;
                            }

                            if (y + rowHeight > e.MarginBounds.Bottom)
                            {
                                e.HasMorePages = true;
                                currentRowIndex = tempRowIndex;
                                return;
                            }

                            foreach (var col in colsToPrint)
                            {
                                int cellWidth = getCellWidth(col);
                                string cellValue = "";

                                object rawValue = row.Cells[col.Index].Value;
                                if (rawValue != DBNull.Value && rawValue != null)
                                {
                                    if (rawValue is DateTime dt)
                                    {
                                        if (col.Name.Equals("Ora", StringComparison.OrdinalIgnoreCase) || col.HeaderText.StartsWith("Ora", StringComparison.OrdinalIgnoreCase))
                                            cellValue = dt.ToString("HH:mm:ss");
                                        else if (col.Name.Equals("Data", StringComparison.OrdinalIgnoreCase) || col.HeaderText.StartsWith("Data", StringComparison.OrdinalIgnoreCase))
                                            cellValue = dt.ToString("yyyy-MM-dd");
                                        else
                                            cellValue = dt.ToString("dd.MM.yyyy HH:mm:ss");
                                    }
                                    else
                                    {
                                        cellValue = rawValue.ToString();
                                    }
                                }

                                e.Graphics.DrawRectangle(Pens.LightGray, x, y, cellWidth, rowHeight);
                                e.Graphics.DrawString(cellValue, fontBody, Brushes.Black, new RectangleF(x + 3, y + 2, cellWidth - 5, rowHeight - 4));
                                x += cellWidth;
                            }

                            y += rowHeight;
                            x = e.MarginBounds.Left;
                            tempRowIndex++;
                        }

                        if (currentGroupIndex < columnGroups.Count - 1)
                        {
                            currentGroupIndex++;
                            currentRowIndex = 0;

                            y += spacing;

                            if ((headerHeight + singleBlockDataHeight) >= halfPageHeight)
                            {
                                e.HasMorePages = true;
                                return;
                            }

                            if (y + totalSingleBlockHeight < e.MarginBounds.Bottom)
                            {
                                continue;
                            }
                            else
                            {
                                e.HasMorePages = true;
                                return;
                            }
                        }
                        else
                        {
                            e.HasMorePages = false;
                            break;
                        }
                    }
                };

                using (PrintPreviewDialog previewDlg = new PrintPreviewDialog())
                {
                    previewDlg.Document = printDoc;
                    previewDlg.WindowState = FormWindowState.Maximized;
                    previewDlg.ShowDialog();
                }
            }
        }
        #endregion
        #region Действия с полями и диагностика

        private void LoadBoundaryValues()
        {
            string dbPath = GetAvailableDatabasePath();
            if (!File.Exists(dbPath)) return;

            string connString = GetConnectionString();
            if (string.IsNullOrEmpty(connString)) return;

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.Open();

                    string dateQuery = "SELECT MIN([Data]) as FirstDate, MAX([Data]) as LastDate FROM [Process_DataE]";
                    using (OleDbCommand cmd = new OleDbCommand(dateQuery, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            firstDate = reader["FirstDate"] != DBNull.Value ? Convert.ToDateTime(reader["FirstDate"]) : DateTime.Now.AddDays(-7);
                            lastDate = reader["LastDate"] != DBNull.Value ? Convert.ToDateTime(reader["LastDate"]) : DateTime.Now;
                            lastDateMinus7Days = lastDate.AddDays(-7);
                        }
                    }

                    string timeQuery = "SELECT MIN([Ora]) as FirstTime, MAX([Ora]) as LastTime FROM [Process_DataE]";
                    using (OleDbCommand cmd = new OleDbCommand(timeQuery, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (reader["FirstTime"] != DBNull.Value)
                            {
                                DateTime dt = Convert.ToDateTime(reader["FirstTime"]);
                                firstTime = dt.ToString("HH:mm");
                            }
                            else firstTime = "00:00";

                            if (reader["LastTime"] != DBNull.Value)
                            {
                                DateTime dt = Convert.ToDateTime(reader["LastTime"]);
                                lastTime = dt.ToString("HH:mm");
                            }
                            else lastTime = "23:59";
                        }
                    }

                    string cycleQuery = "SELECT MIN([Cy]) as FirstCycle, MAX([Cy]) as LastCycle FROM [Process_DataE]";
                    using (OleDbCommand cmd = new OleDbCommand(cycleQuery, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            firstCycle = reader["FirstCycle"] != DBNull.Value ? Convert.ToInt32(reader["FirstCycle"]) : 1;
                            lastCycle = reader["LastCycle"] != DBNull.Value ? Convert.ToInt32(reader["LastCycle"]) : 1000;
                        }
                    }
                }

                UpdateButtonTexts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке граничных значений:\n{ex.Message}", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetDefaultBoundaryValues();
                UpdateButtonTexts();
            }
        }

        private void SetDefaultBoundaryValues()
        {
            firstDate = DateTime.Now.AddDays(-7);
            lastDate = DateTime.Now;
            lastDateMinus7Days = DateTime.Now.AddDays(-7);
            firstTime = "00:00";
            lastTime = "23:59";
            firstCycle = 1;
            lastCycle = 1000;
        }

        private void UpdateButtonTexts()
        {
            btnFirstDateODR.Text = $"{firstDate:yyyy-MM-dd}";
            btnLastDateODR.Text = $"{lastDate:yyyy-MM-dd}";
            btnFirstTimeODR.Text = $"{firstTime}";
            btnLastTimeODR.Text = $"{lastTime}";
            btnFirstCycleODR.Text = $"{firstCycle}";
            btnLastCycleODR.Text = $"{lastCycle}";
            btn7dDateODR.Text = $"{lastDateMinus7Days:yyyy-MM-dd}";

            btnFirstDatePDR.Text = $"{firstDate:yyyy-MM-dd}";
            btnLastDatePDR.Text = $"{lastDate:yyyy-MM-dd}";
            btnFirstTimePDR.Text = $"{firstTime}";
            btnLastTimePDR.Text = $"{lastTime}";
            btn7dDatePDR.Text = $"{lastDateMinus7Days:yyyy-MM-dd}";
        }

        private void SetFirstLastDate(TextBox targetBox, bool isFirst)
        {
            string newValue = isFirst ? firstDate.ToString("yyyy-MM-dd") : lastDate.ToString("yyyy-MM-dd");
            targetBox.ForeColor = Color.Black;
            targetBox.Text = newValue;
        }

        private void SetFirstLastTime(TextBox targetBox, bool isFirst)
        {
            string newValue = isFirst ? firstTime : lastTime;
            targetBox.ForeColor = Color.Black;
            targetBox.Text = newValue;
        }

        private void SetFirstLastCycle(TextBox targetBox, bool isFirst)
        {
            string newValue = isFirst ? firstCycle.ToString() : lastCycle.ToString();
            targetBox.ForeColor = Color.Black;
            targetBox.Text = newValue;
        }

        private void SetDateMinus7DaysToField(TextBox targetBox)
        {
            string newValue = lastDateMinus7Days.ToString("yyyy-MM-dd");
            targetBox.ForeColor = Color.Black;
            targetBox.Text = newValue;
        }

        private void BtnDiagnose_Click(object sender, EventArgs e)
        {
            string dbPath = GetAvailableDatabasePath();
            if (!File.Exists(dbPath))
            {
                MessageBox.Show($"Файл не найден: {dbPath}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine("ДИАГНОСТИКА БАЗЫ ДАННЫХ");
            sb.AppendLine($"Файл: {Path.GetFileName(dbPath)}");
            sb.AppendLine($"Путь: {dbPath}");
            sb.AppendLine($"Размер: {new FileInfo(dbPath).Length / 1024} KB");
            sb.AppendLine();

            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine("ИНФОРМАЦИЯ ИЗ INF ФАЙЛА");
            sb.AppendLine($"PROCESS_TABLE: {InfParser.ProcessTableName}");
            sb.AppendLine($"PRODUCT_TABLE: {InfParser.ProductTableName}");
            sb.AppendLine();

            // GRID настройки
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine("GRID НАСТРОЙКИ:");
            sb.AppendLine($"  Process полей: {infParser.GridProcessFields.Count}");
            sb.AppendLine($"  Product полей: {infParser.GridProductFields.Count}");

            // GRID Process (первые 20)
            if (infParser.GridProcessFields.Count > 0)
            {
                sb.AppendLine("  GRID Process (первые 20):");
                for (int i = 0; i < Math.Min(20, infParser.GridProcessFields.Count); i++)
                {
                    sb.AppendLine($"    [{i + 1}] {infParser.GridProcessFields[i]}");
                }
                if (infParser.GridProcessFields.Count > 20)
                    sb.AppendLine($"    ... и еще {infParser.GridProcessFields.Count - 20} полей");
            }

            // GRID Product
            if (infParser.GridProductFields.Count > 0)
            {
                sb.AppendLine("  GRID Product:");
                foreach (var field in infParser.GridProductFields)
                {
                    sb.AppendLine($"    • {field}");
                }
            }

            // Отчёты
            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine("ОТЧЁТЫ:");
            sb.AppendLine($"  PROCESS_REPORT1: {infParser.ProcessReportFields.Count} строк");
            foreach (var reportField in infParser.ProcessReportFields)
            {
                sb.AppendLine($"    • {reportField.Name} [{reportField.Unit}]: {reportField.Formula}");
            }

            sb.AppendLine($"  PRODUCT_REPORT1: {infParser.ProductReportFields.Count} строк");
            foreach (var reportField in infParser.ProductReportFields)
            {
                sb.AppendLine($"    • {reportField.Name} [{reportField.Unit}]: {reportField.Formula}");
            }

            // PROCESS_TABLE поля (только валидные)
            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            int validProcessCount = infParser.ProcessFields.Count;
            sb.AppendLine($"PROCESS_TABLE валидные поля (всего: {validProcessCount} из 100)");

            if (validProcessCount > 0)
            {
                sb.AppendLine("  Список валидных полей:");
                int displayCount = 0;
                foreach (var field in infParser.ProcessFields)
                {
                    if (displayCount++ >= 40)
                    {
                        sb.AppendLine($"    ... и еще {validProcessCount - 40} валидных полей");
                        break;
                    }
                    sb.AppendLine($"    • {field.Name} - {field.Description} [{field.Unit}] (адрес: {field.PlcAddress})");
                }
            }
            else
            {
                sb.AppendLine("  Нет валидных полей (все поля имеют PlcAddress = null)");
            }

            // PRODUCT_TABLE поля (только с формулами)
            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            var validProductFields = infParser.ProductFields.Where(f => !string.IsNullOrEmpty(f.Formula) && f.Formula != "null").ToList();
            sb.AppendLine($"PRODUCT_TABLE поля с формулами (всего: {validProductFields.Count} из {infParser.ProductFields.Count})");
            foreach (var field in validProductFields)
            {
                sb.AppendLine($"    • {field.Name} - {field.Description} [{field.Unit}]");
                sb.AppendLine($"      Формула: {field.Formula}");
            }

            // Дополнительно
            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine("ДОПОЛНИТЕЛЬНО:");
            sb.AppendLine($"  FileDati: {infParser.FileDati}");
            sb.AppendLine($"  FileSettimanale: {infParser.FileSettimanale}");
            sb.AppendLine($"  NomeFileCopia: {infParser.NomeFileCopia ?? "(не задан)"}");
            sb.AppendLine($"  MinMaxFields: {infParser.MinMaxFields.Count} полей");
            if (infParser.MinMaxFields.Count > 0)
            {
                sb.AppendLine($"    {string.Join(", ", infParser.MinMaxFields)}");
            }

            // Информация о подключении
            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine("ПОДКЛЮЧЕНИЕ К БАЗЕ:");
            try
            {
                string connString = GetConnectionString();
                if (!string.IsNullOrEmpty(connString))
                {
                    sb.AppendLine($"  Драйвер: {(connString.Contains("ACE") ? "ACE OLEDB" : "Jet OLEDB")}");
                    using (var testConn = new OleDbConnection(connString))
                    {
                        testConn.Open();
                        sb.AppendLine("  Статус: ✅ Подключение успешно");
                        sb.AppendLine($"  Версия базы: {testConn.ServerVersion}");
                    }
                }
                else
                {
                    sb.AppendLine("  Статус: ❌ Не удалось получить строку подключения");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  Статус: ❌ Ошибка подключения");
                sb.AppendLine($"  Ошибка: {ex.Message}");
            }

            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");

            // Показываем результат в отдельном окне с возможностью копирования
            Form diagnosticsForm = new Form();
            diagnosticsForm.Text = "Диагностика INF и базы данных";
            diagnosticsForm.Size = new System.Drawing.Size(900, 700);
            diagnosticsForm.StartPosition = FormStartPosition.CenterScreen;

            TextBox textBox = new TextBox();
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Both;
            textBox.Dock = DockStyle.Fill;
            textBox.Font = new System.Drawing.Font("Consolas", 9);
            textBox.Text = sb.ToString();
            textBox.ReadOnly = true;

            diagnosticsForm.Controls.Add(textBox);
            diagnosticsForm.ShowDialog();
        }
        #endregion

    }
}