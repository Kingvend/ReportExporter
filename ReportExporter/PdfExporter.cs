using System.IO;

public static class PdfExporter
{
	public static void Export(string pdfReportPath, string imagePath, string title)
	{
		// Сначала создаем временный docx файл
		string tempDocxPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");

		try
		{
			// Создаем DOCX документ с помощью основного класса
			PngToWordExporter.Export(tempDocxPath, imagePath, title);

			// Конвертируем DOCX в PDF
			// ConvertDocxToPdf(tempDocxPath, pdfReportPath);
			DevExpressExportHelper.DevExpressExportHelper.ConvertDocxToPdf(tempDocxPath, pdfReportPath);
		}
		finally
		{
			// Удаляем временный файл
			if (File.Exists(tempDocxPath))
			{
				try
				{
					File.Delete(tempDocxPath);
				}
				catch
				{
					// Игнорируем ошибки удаления временного файла
				}
			}
		}
	}
}