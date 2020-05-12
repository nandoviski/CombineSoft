using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CombineSoft
{
	public class TimeCourseExtractorExcelUtil
	{
		Color excelColor = Color.LightBlue;
		string currentGroup = null;

		const int CellBox = 1;
		const int CellGroup = 2;
		const int CellSubject = 3;
		const int CellActive = 4;
		const int CellInactive = 5;
		const int CellInfusions = 6;
		const int CellTotalActivity = 7;

		List<string> ratsWithSalina = new List<string>
		{
			"8F", "10F", "8M", "12M", "21F", "26F", "24M", "27M", "41F", "44F", "42M", "45M", "57F", "61F", "56M", "60M"
		};

		public bool CreateExcel(List<TimeCourseExtractor> fileDatas, string filePath)
		{
			try
			{
				ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

				using (var excelPackage = new ExcelPackage(new MemoryStream()))
				{
					var ws1 = excelPackage.Workbook.Worksheets.Add("TimeCourse");
					var columnMultipleir = 1;

					foreach (var file in fileDatas)
					{
						CreateCellTitles(ws1, file.Subject, columnMultipleir);

						foreach (var item in file.TimeCountPerAction)
						{
							var row = 3;
							var column = 3;
							if (item.Action == "E")
							{
								column = 3 + columnMultipleir;
							}
							else if (item.Action == "F")
							{
								column = 4 + columnMultipleir;
							}
							else if (item.Action == "G")
							{
								column = 5 + columnMultipleir;
							}
							else if (item.Action == "H")
							{
								column = 6 + columnMultipleir;
							}

							foreach (var aaa in item.Times)
							{
								ws1.Cells[row++, column].Value = aaa.Value;
							}

							ws1.Cells[row++, column].Value = item.CalculateTotal();
						}

						columnMultipleir+=7;
					}

					excelPackage.SaveAs(new FileInfo(filePath));
				}

				return true;
			}
			catch (Exception ex)
			{
				var msg = ex.Message;
				if (ex.InnerException != null)
				{
					msg += "\n\nInnerException: " + ex.InnerException.Message;
				}
				MessageBox.Show(msg, "Error Generating Excel", MessageBoxButton.OK, MessageBoxImage.Error);
				return false;
			}
		}

		void CreateCellTitles(ExcelWorksheet ws1, string subject, int columnMultipleir)
		{
			var row = 1;

			ws1.Cells[row, 1 + columnMultipleir].Value = "Rat " + subject;
			ws1.Cells[row, 1 + columnMultipleir].Style.Font.Bold = true;
			ws1.Cells[row, 1 + columnMultipleir].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			ws1.Cells[row, 1 + columnMultipleir, row, 6 + columnMultipleir].Merge = true;

			row++;
			ws1.Cells[row, 1 + columnMultipleir].Value = "secs";
			ws1.Cells[row, 1 + columnMultipleir].Style.Font.Bold = true;
			ws1.Cells[row, 2 + columnMultipleir].Value = "mins";
			ws1.Cells[row, 2 + columnMultipleir].Style.Font.Bold = true;
			ws1.Cells[row, 3 + columnMultipleir].Value = "E = Active";
			ws1.Cells[row, 3 + columnMultipleir].Style.Font.Bold = true;
			ws1.Cells[row, 4 + columnMultipleir].Value = "F = Inactive";
			ws1.Cells[row, 4 + columnMultipleir].Style.Font.Bold = true;
			ws1.Cells[row, 5 + columnMultipleir].Value = "G = Infusions";
			ws1.Cells[row, 5 + columnMultipleir].Style.Font.Bold = true;
			ws1.Cells[row, 6 + columnMultipleir].Value = "H = Loco";
			ws1.Cells[row, 6 + columnMultipleir].Style.Font.Bold = true;

			for (int i = 1; i <= TimeCourseExtractor.MultiplierCount; i++)
			{
				var totalInSeconds = i * TimeCourseExtractor.TimeInSeconds;
				ws1.Cells[(row + i), 1 + columnMultipleir].Value = totalInSeconds;
				ws1.Cells[(row + i), 2 + columnMultipleir].Value = totalInSeconds / 60;
			}
		}

		void CellStyle(ExcelRange cell, Color color, bool alignmentCenter = true, bool border = true)
		{
			cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
			cell.Style.Fill.BackgroundColor.SetColor(color);
			if (alignmentCenter)
			{
				cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			}
			if (border)
			{
				cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
			}
		}
	}
}
