using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.IO;

public static class PngToWordExporter
{
	public static void Export(string reportPath, string imagePath, string title)
	{
		using (var wordDocument = WordprocessingDocument.Create(reportPath, WordprocessingDocumentType.Document))
		{
			// Создаем главную часть документа
			MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
			mainPart.Document = new Document();
			mainPart.Document.AppendChild(new Body());

			// Настраиваем параметры страницы
			var sectionProps = CreateSectionProperties();
			mainPart.Document.Body.Append(sectionProps);

			// Добавляем заголовок
			AddTitle(mainPart, title);

			// Добавляем изображение
			AddImage(mainPart, imagePath);

			mainPart.Document.Save();
		}
	}

	private static SectionProperties CreateSectionProperties()
	{
		// 1 см = 567 TWIP
		uint marginSize = 567;

		return new SectionProperties(
			new PageSize()
			{
				Width = 16840U,    // Ширина A4 в альбомной ориентации (в TWIP)
				Height = 11900U,   // Высота A4 в альбомной ориентации (в TWIP)
				Orient = PageOrientationValues.Landscape
			},
			new PageMargin()       // Устанавливаем поля по 1 см
			{
				Top = (Int32Value)(int)marginSize,
				Right = (UInt32Value)marginSize,
				Bottom = (Int32Value)(int)marginSize,
				Left = (UInt32Value)marginSize,
				Header = (UInt32Value)0U,
				Footer = (UInt32Value)0U
			}
		);
	}

	private static void AddTitle(MainDocumentPart mainPart, string titleText)
	{
		var titleParagraph = new Paragraph(
			new ParagraphProperties(
				new Justification() { Val = JustificationValues.Center },
				new SpacingBetweenLines() { After = "200" } // Отступ после заголовка
			),
			new Run(
				new RunProperties(
					new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma" },
					new FontSize() { Val = "28" },  // 14pt * 2 = 28 (полуточек)
					new Color() { Val = "000000" }
				),
				new Text(titleText)
			)
		);

		mainPart.Document.Body.Append(titleParagraph);
	}

	private static void AddImage(MainDocumentPart mainPart, string imagePath)
	{
		// Создаем часть изображения
		ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
		using (FileStream stream = new FileStream(imagePath, FileMode.Open))
		{
			imagePart.FeedData(stream);
		}

		// Получаем размеры изображения
		using (System.Drawing.Image image = System.Drawing.Image.FromFile(imagePath))
		{
			// Рассчитываем размеры для заполнения ширины страницы с учетом полей
			// Доступная ширина = ширина страницы - левое поле - правое поле
			// 16840 TWIP (ширина страницы) - 567 * 2 (поля) = 15706 TWIP
			long availableWidthInTwip = 16840L - 567L * 2;

			// Конвертируем доступную ширину в EMU
			long availableWidthInEmu = availableWidthInTwip * 635;

			// Рассчитываем размеры изображения в EMU
			long imageWidthInEmu = (long)(image.Width * 9525);
			long imageHeightInEmu = (long)(image.Height * 9525);

			// Вычисляем новую высоту с сохранением пропорций
			long newHeight = imageHeightInEmu * availableWidthInEmu / imageWidthInEmu;

			// Создаем элемент Drawing с Anchor для позиционирования
			var element = CreateImageElementWithAnchor(
				mainPart.GetIdOfPart(imagePart),
				"Exported Image",
				availableWidthInEmu,
				newHeight
			);

			// Просто добавляем элемент в абзац без дополнительного форматирования
			var imageParagraph = new Paragraph(new Run(element));
			mainPart.Document.Body.Append(imageParagraph);
		}
	}

	private static Drawing CreateImageElementWithAnchor(string relationshipId, string title, long width, long height)
	{
		// Вычисляем позицию для центрирования по горизонтали
		// Ширина страницы в EMU: 16840 * 635 = 10693400
		// Отступы: 567 * 635 * 2 = 720090
		// Доступная ширина: 10693400 - 720090 = 9967310
		// Центрирование: (10693400 - width) / 2
		long pageWidthInEmu = 16840L * 635L;
		long horizontalOffset = (pageWidthInEmu - width) / 2L;

		// Вычисляем позицию для центрирования по вертикали
		// Высота страницы в EMU: 11900 * 635 = 7556500
		// Отступы: 567 * 635 * 2 = 720090
		// Доступная высота: 7556500 - 720090 = 6836410
		// Центрирование: (7556500 - height) / 2
		long pageHeightInEmu = 11900L * 635L;
		long verticalOffset = (pageHeightInEmu - height) / 2L;

		return new Drawing(
			new DW.Anchor(
				new DW.SimplePosition() { X = 0L, Y = 0L },
				new DW.HorizontalPosition(
					new DW.PositionOffset() { Text = horizontalOffset.ToString() }
				)
				{
					RelativeFrom = DW.HorizontalRelativePositionValues.Page
				},
				new DW.VerticalPosition(
					new DW.PositionOffset() { Text = verticalOffset.ToString() }
				)
				{
					RelativeFrom = DW.VerticalRelativePositionValues.Page
				},
				new DW.Extent() { Cx = width, Cy = height },
				new DW.EffectExtent()
				{
					LeftEdge = 0L,
					TopEdge = 0L,
					RightEdge = 0L,
					BottomEdge = 0L
				},
				new DW.WrapNone(), // Без обтекания текстом
				new DW.DocProperties()
				{
					Id = (UInt32Value)1U,
					Name = title
				},
				new DW.NonVisualGraphicFrameDrawingProperties(
					new A.GraphicFrameLocks() { NoChangeAspect = true }
				),
				new A.Graphic(
					new A.GraphicData(
						new A.Pictures.Picture(
							new A.Pictures.NonVisualPictureProperties(
								new A.Pictures.NonVisualDrawingProperties()
								{
									Id = (UInt32Value)0U,
									Name = title
								},
								new A.Pictures.NonVisualPictureDrawingProperties()
							),
							new A.Pictures.BlipFill(
								new A.Blip()
								{
									Embed = relationshipId,
									CompressionState = A.BlipCompressionValues.Print
								},
								new A.Stretch(
									new A.FillRectangle()
								)
							),
							new A.Pictures.ShapeProperties(
								new A.Transform2D(
									new A.Offset() { X = 0L, Y = 0L },
									new A.Extents() { Cx = width, Cy = height }
								),
								new A.PresetGeometry(
									new A.AdjustValueList()
								)
								{ Preset = A.ShapeTypeValues.Rectangle }
							)
						)
					)
					{ Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
				)
			)
			{
				DistanceFromTop = (UInt32Value)0U,
				DistanceFromBottom = (UInt32Value)0U,
				DistanceFromLeft = (UInt32Value)0U,
				DistanceFromRight = (UInt32Value)0U,
				SimplePos = false,
				RelativeHeight = (UInt32Value)2516608U, // Помещаем изображение за текстом
				BehindDoc = false,
				Locked = false,
				LayoutInCell = true,
				AllowOverlap = true
			}
		);
	}
}