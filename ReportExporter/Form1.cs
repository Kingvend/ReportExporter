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


	public partial class Form1 : Form
	{
		public string ImagePath { get; set; } = null;
		public string ReportPath { get; set; } = null;
		public Form1()
		{
			InitializeComponent();
		}

		private void simpleButton1_Click(object sender, EventArgs e)
		{
			string reportPath = Path.Combine(ReportPath, "report.docx");

			PngToWordExporter.Export(reportPath, ImagePath, "Сводный отчет по субъекту РФ Ханты-Мансийский АО");
		}

		private void simpleButton2_Click(object sender, EventArgs e)
		{

		}

		private void simpleButton3_Click(object sender, EventArgs e)
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
				}
			}
		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}

		private void simpleButton4_Click(object sender, EventArgs e)
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
					textEdit1.Text = selectedPath;
				}
			}
		}
	}
}
