using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace K450DataExporter
{
    public partial class Form2 : Form
    {
        public string SelectedInfPath { get; private set; }

        public Form2()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            LoadInfFiles();

            // Привязываем обработчики к кнопкам
            btnOpen.Click += btnOpen_Click;
            btnCancel.Click += btnCancel_Click;

            // ПРИВЯЗЫВАЕМ ОБРАБОТЧИК ДЛЯ LISTBOX (это важно!)
            listBoxInfFiles.SelectedIndexChanged += listBoxInfFiles_SelectedIndexChanged;

            // Изначально кнопка OK неактивна
            btnOpen.Enabled = false;
        }

        private void LoadInfFiles()
        {
            string exeFolder = Application.StartupPath;
            string[] infFiles = Directory.GetFiles(exeFolder, "*.inf");

            listBoxInfFiles.Items.Clear();
            foreach (string file in infFiles)
            {
                listBoxInfFiles.Items.Add(Path.GetFileName(file));
            }

            if (infFiles.Length == 0)
            {
                MessageBox.Show("Не найдено INF-файлов в папке:\n" + exeFolder,
                                "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btnOpen.Enabled = false;
            }
        }

        private void listBoxInfFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Кнопка OK активна только когда выбран элемент в списке
            btnOpen.Enabled = (listBoxInfFiles.SelectedItem != null);
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (listBoxInfFiles.SelectedItem == null)
            {
                return;
            }

            string selectedFile = listBoxInfFiles.SelectedItem.ToString();
            SelectedInfPath = Path.Combine(Application.StartupPath, selectedFile);

            if (!File.Exists(SelectedInfPath))
            {
                MessageBox.Show($"Файл не найден:\n{SelectedInfPath}", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}