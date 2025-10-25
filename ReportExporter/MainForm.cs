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

namespace ReportExporter
{


	public partial class MainForm : Form
	{
		public string ImagePath { get; set; } = null;
		public string ReportPath { get; set; } = null;
		public MainForm()
		{
			InitializeComponent();
		}

		private void BtnWordExport_Click(object sender, EventArgs e)
		{
			try
			{
				string reportPath = Path.Combine(ReportPath, "report.docx");
				PngToWordExporter.Export(reportPath, ImagePath, "Сводный отчет по субъекту РФ Ханты-Мансийский АО");

				MessageBox.Show("Отчет сохранен успешно");
			}
			catch (Exception)
			{
				MessageBox.Show("Ошибка");
			}
		}

		private void BtnPdfExport_Click(object sender, EventArgs e)
		{
			try
			{
				string reportPath = Path.Combine(ReportPath, "report.pdf");
				PdfExporter.Export(reportPath, ImagePath, "Сводный отчет по субъекту РФ Ханты-Мансийский АО");

				MessageBox.Show("Отчет сохранен успешно");
			}
			catch (Exception)
			{
				MessageBox.Show("Ошибка");
			}
		}

		private void BtnOpenImage_Click(object sender, EventArgs e)
		{
			using (OpenFileDialog openFileDialog = new OpenFileDialog())
			{
				openFileDialog.InitialDirectory = "c:\\";
				openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
				openFileDialog.FilterIndex = 2;
				openFileDialog.RestoreDirectory = true;

				if (openFileDialog.ShowDialog() == DialogResult.OK)
				{
					//Get the path of specified file
					ImagePath = openFileDialog.FileName;
					textBox1.Text = ImagePath;
				}
			}
		}

		private void BtnOpenReportFolder_Click(object sender, EventArgs e)
		{
			using (var folderDialog = new FolderBrowserDialog())
			{
				// Настройки диалога
				folderDialog.Description = "Выберите папку"; // Текст над деревом папок
				folderDialog.RootFolder = Environment.SpecialFolder.MyComputer; // Начальная папка
				folderDialog.ShowNewFolderButton = true; // Разрешить создание новых папок

				// Открыть диалог и проверить результат
				if (folderDialog.ShowDialog() == DialogResult.OK)
				{
					string selectedPath = folderDialog.SelectedPath;
					// Используйте selectedPath (например, выведите в TextBox)
					ReportPath = selectedPath;
					textBox2.Text = selectedPath;
				}
			}
		}
	}
}
