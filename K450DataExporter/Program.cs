using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace K450DataExporter
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string infPath = null;

            // Проверяем параметры командной строки
            foreach (string arg in args)
            {
                if (arg.StartsWith("/F", StringComparison.OrdinalIgnoreCase))
                {
                    infPath = arg.Substring(2).Trim('"');
                    break;
                }
                else if (arg.StartsWith("-f", StringComparison.OrdinalIgnoreCase))
                {
                    infPath = arg.Substring(2).Trim('"');
                    break;
                }
            }

            // Если параметр передан
            if (!string.IsNullOrEmpty(infPath))
            {
                if (!File.Exists(infPath))
                {
                    MessageBox.Show($"INF-файл не найден:\n{infPath}", "Ошибка",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Form1 mainForm = new Form1(infPath);
                Application.Run(mainForm);
            }
            else
            {
                // Без параметра — показываем диалог выбора
                Form2 dialog = new Form2();
                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(dialog.SelectedInfPath))
                {
                    Form1 mainForm = new Form1(dialog.SelectedInfPath);
                    Application.Run(mainForm);
                }
            }
        }
    }
}