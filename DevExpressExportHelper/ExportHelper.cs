using DevExpress.XtraPrinting;
using DevExpress.XtraRichEdit;
using System;
using System.IO;
using System.Linq;

namespace DevExpressExportHelper
{
	public static class DevExpressExportHelper
	{
		public static void ConvertDocxToPdf(string docxPath, string pdfPath)
		{
			if (!File.Exists(docxPath))
				throw new FileNotFoundException($"DOCX файл не найден: {docxPath}");

			// Создаем директорию для PDF если не существует
			string pdfDirectory = Path.GetDirectoryName(pdfPath);
			if (!Directory.Exists(pdfDirectory))
				Directory.CreateDirectory(pdfDirectory);

			// Используем RichEditControl для загрузки DOCX и экспорта в PDF
			using (RichEditDocumentServer documentServer = new RichEditDocumentServer())
			{
				try
				{
					// Загружаем DOCX документ
					documentServer.LoadDocument(docxPath, DocumentFormat.OpenXml);

					// Настраиваем опции экспорта в PDF
					PdfExportOptions pdfOptions = new PdfExportOptions()
					{
						/* Базовые настройки PDF*/
						Compressed = true,
						ImageQuality = PdfJpegImageQuality.High
					};

					// Настройки документа
					pdfOptions.DocumentOptions.Application = "Report Generator";
					pdfOptions.DocumentOptions.Title = "Generated Report";
					pdfOptions.DocumentOptions.Author = "Report System";

					// Экспортируем в PDF
					documentServer.ExportToPdf(pdfPath, pdfOptions);
				}
				catch (Exception ex)
				{
					throw new Exception($"Ошибка конвертации DevExpress: {ex.Message}", ex);
				}
			}
		}
	}
}
