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
		public Form1()
		{
			InitializeComponent();
		}

		private void simpleButton1_Click(object sender, EventArgs e)
		{
			string reportPath = Path.Combine(Path.GetTempPath(), "report.docx");
			// string imagePath = @"C:\Users\Alex\Desktop\mapImage.png";
			string imagePath = @"C:\Users\Alex\Desktop\123.png";

			PngToWordExporter.Export(reportPath, imagePath, "Сводный отчет по субъекту РФ Ханты-Мансийский АО");
		}

		private void simpleButton2_Click(object sender, EventArgs e)
		{

		}

		private void simpleButton3_Click(object sender, EventArgs e)
		{
			
		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}
	}
}
