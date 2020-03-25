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
	public class ExcelUtil
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

		public bool CreateExcel(List<FileData> fileDatas, string filePath)
		{
			try
			{
				ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

				using (var excelPackage = new ExcelPackage(new MemoryStream()))
				{
					var ws1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
					
					CreateCellTitles(ws1);

					var filesOrdered = fileDatas.OrderBy(f => f.StartDate).ThenBy(f => f.Group).ThenBy(f => f.Gender).ThenBy(f => f.RatNumber);
					var row = 2;
					var rowNoSalina = 2;

					foreach (var grp in new[] { "H", "S", "0", "A" })
					{
						string group = null;
						string gender = null;
						var color = GetColor(grp);
						var highlightColor = Color.FromArgb(146, 208, 80);
						var filteredByGroup = filesOrdered.Where(f => f.Group.ToUpper() == grp);

						var averageFormulaActive = new List<string>();
						var averageFormulaInactive = new List<string>();
						var averageFormulaInfusions = new List<string>();
						var averageFormulaTotalActivity = new List<string>();

						var salinaRats = false;
						var shiftCells = 9;
						var averageNoSalinaFormulaActive = new List<string>();
						var averageNoSalinaFormulaInactive = new List<string>();
						var averageNoSalinaFormulaInfusions = new List<string>();
						var averageNoSalinaFormulaTotalActivity = new List<string>();

						foreach (var item in filteredByGroup)
						{
							salinaRats = ratsWithSalina.Contains(item.Subject);

							if (!string.IsNullOrEmpty(group))
							{
								if (group != item.Group || gender != item.Gender)
								{
									CreateEmptyLine(ws1, row, color, averageFormulaActive, averageFormulaInactive, averageFormulaInfusions, averageFormulaTotalActivity);
									CreateEmptyLine(ws1, rowNoSalina, color, averageNoSalinaFormulaActive, averageNoSalinaFormulaInactive, averageNoSalinaFormulaInfusions, averageNoSalinaFormulaTotalActivity, shiftCells);
									row += 3;
									rowNoSalina += 3;

									averageFormulaActive = new List<string>();
									averageFormulaInactive = new List<string>();
									averageFormulaInfusions = new List<string>();
									averageFormulaTotalActivity = new List<string>();

									averageNoSalinaFormulaActive = new List<string>();
									averageNoSalinaFormulaInactive = new List<string>();
									averageNoSalinaFormulaInfusions = new List<string>();
									averageNoSalinaFormulaTotalActivity = new List<string>();
								}
							}

							var currentColor = color;

							if (salinaRats)
							{
								currentColor = highlightColor;
							}
							else
							{
								averageFormulaActive.Add(ws1.Cells[row, CellActive].Address);
								averageFormulaInactive.Add(ws1.Cells[row, CellInactive].Address);
								averageFormulaInfusions.Add(ws1.Cells[row, CellInfusions].Address);
								averageFormulaTotalActivity.Add(ws1.Cells[row, CellTotalActivity].Address);
							}

							ws1.Cells[row, CellBox].Value = item.Box;
							CellStyle(ws1.Cells[row, CellBox], currentColor);

							ws1.Cells[row, CellGroup].Value = GroupName(item.Group);
							CellStyle(ws1.Cells[row, CellGroup], currentColor, false);

							ws1.Cells[row, CellSubject].Value = item.Subject;
							CellStyle(ws1.Cells[row, CellSubject], currentColor);

							ws1.Cells[row, CellActive].Value = item.Active;
							CellStyle(ws1.Cells[row, CellActive], currentColor);

							ws1.Cells[row, CellInactive].Value = item.Inactive;
							CellStyle(ws1.Cells[row, CellInactive], currentColor);

							ws1.Cells[row, CellInfusions].Value = item.Infusions;
							CellStyle(ws1.Cells[row, CellInfusions], currentColor);

							ws1.Cells[row, CellTotalActivity].Value = item.TotalActivity;
							CellStyle(ws1.Cells[row, CellTotalActivity], currentColor);

							row++;

							if (!salinaRats)
							{
								ws1.Cells[rowNoSalina, CellBox + shiftCells].Value = item.Box;
								CellStyle(ws1.Cells[rowNoSalina, CellBox + shiftCells], currentColor);

								ws1.Cells[rowNoSalina, CellGroup + shiftCells].Value = GroupName(item.Group);
								CellStyle(ws1.Cells[rowNoSalina, CellGroup + shiftCells], currentColor, false);

								ws1.Cells[rowNoSalina, CellSubject + shiftCells].Value = item.Subject;
								CellStyle(ws1.Cells[rowNoSalina, CellSubject + shiftCells], currentColor);

								ws1.Cells[rowNoSalina, CellActive + shiftCells].Value = item.Active;
								CellStyle(ws1.Cells[rowNoSalina, CellActive + shiftCells], currentColor);

								ws1.Cells[rowNoSalina, CellInactive + shiftCells].Value = item.Inactive;
								CellStyle(ws1.Cells[rowNoSalina, CellInactive + shiftCells], currentColor);

								ws1.Cells[rowNoSalina, CellInfusions + shiftCells].Value = item.Infusions;
								CellStyle(ws1.Cells[rowNoSalina, CellInfusions + shiftCells], currentColor);

								ws1.Cells[rowNoSalina, CellTotalActivity + shiftCells].Value = item.TotalActivity;
								CellStyle(ws1.Cells[rowNoSalina, CellTotalActivity + shiftCells], currentColor);

								averageNoSalinaFormulaActive.Add(ws1.Cells[rowNoSalina, CellActive + shiftCells].Address);
								averageNoSalinaFormulaInactive.Add(ws1.Cells[rowNoSalina, CellInactive + shiftCells].Address);
								averageNoSalinaFormulaInfusions.Add(ws1.Cells[rowNoSalina, CellInfusions + shiftCells].Address);
								averageNoSalinaFormulaTotalActivity.Add(ws1.Cells[rowNoSalina, CellTotalActivity + shiftCells].Address);

								rowNoSalina++;
							}

							group = item.Group;
							gender = item.Gender;
						}

						if (filteredByGroup.Any())
						{
							CreateEmptyLine(ws1, row, color, averageFormulaActive, averageFormulaInactive, averageFormulaInfusions, averageFormulaTotalActivity);
							CreateEmptyLine(ws1, rowNoSalina, color, averageNoSalinaFormulaActive, averageNoSalinaFormulaInactive, averageNoSalinaFormulaInfusions, averageNoSalinaFormulaTotalActivity, shiftCells);
							row += 3;
							rowNoSalina += 3;

							averageFormulaActive = new List<string>();
							averageFormulaInactive = new List<string>();
							averageFormulaInfusions = new List<string>();
							averageFormulaTotalActivity = new List<string>();

							averageNoSalinaFormulaActive = new List<string>();
							averageNoSalinaFormulaInactive = new List<string>();
							averageNoSalinaFormulaInfusions = new List<string>();
							averageNoSalinaFormulaTotalActivity = new List<string>();
						}
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

		void CreateCellTitles(ExcelWorksheet ws1)
		{
			ws1.Cells["A1"].Value = "Chamber";
			ws1.Cells["A1"].Style.Font.Bold = true;
			ws1.Cells["B1"].Value = "Group";
			ws1.Cells["B1"].Style.Font.Bold = true;
			ws1.Cells["C1"].Value = "Rat";
			ws1.Cells["C1"].Style.Font.Bold = true;
			ws1.Cells["D1"].Value = "Active";
			ws1.Cells["D1"].Style.Font.Bold = true;
			ws1.Cells["E1"].Value = "Inactive";
			ws1.Cells["E1"].Style.Font.Bold = true;
			ws1.Cells["F1"].Value = "Infusions";
			ws1.Cells["F1"].Style.Font.Bold = true;
			ws1.Cells["G1"].Value = "Total Activity";
			ws1.Cells["G1"].Style.Font.Bold = true;

			ws1.Cells["J1"].Value = "Chamber";
			ws1.Cells["J1"].Style.Font.Bold = true;
			ws1.Cells["K1"].Value = "Group";
			ws1.Cells["K1"].Style.Font.Bold = true;
			ws1.Cells["L1"].Value = "Rat";
			ws1.Cells["L1"].Style.Font.Bold = true;
			ws1.Cells["M1"].Value = "Active";
			ws1.Cells["M1"].Style.Font.Bold = true;
			ws1.Cells["N1"].Value = "Inactive";
			ws1.Cells["N1"].Style.Font.Bold = true;
			ws1.Cells["O1"].Value = "Infusions";
			ws1.Cells["O1"].Style.Font.Bold = true;
			ws1.Cells["P1"].Value = "Total Activity";
			ws1.Cells["P1"].Style.Font.Bold = true;
		}

		void CreateEmptyLine(ExcelWorksheet ws1, int row, Color color, List<string> averageFormula, List<string> averageFormulaInactive, List<string> averageFormulaInfusions, List<string> averageFormulaTotalActivity, int shiftCells = 0)
		{
			for (int i = 1; i <= 7; i++)
			{
				CellStyle(ws1.Cells[row, i + shiftCells], color, border: false);
				CellStyle(ws1.Cells[row + 1, i + shiftCells], color, border: false);
			}
			ws1.Cells[row, CellSubject + shiftCells].Value = "Mean METH";

			ws1.Cells[row, CellActive + shiftCells].Formula = $"=AVERAGE({string.Join(",", averageFormula)})";
			ws1.Cells[row, CellActive + shiftCells].Style.Numberformat.Format = "0";

			ws1.Cells[row, CellInactive + shiftCells].Formula = $"=AVERAGE({string.Join(",", averageFormulaInactive)})";
			ws1.Cells[row, CellInactive + shiftCells].Style.Numberformat.Format = "0";

			ws1.Cells[row, CellInfusions + shiftCells].Formula = $"=AVERAGE({string.Join(",", averageFormulaInfusions)})";
			ws1.Cells[row, CellInfusions + shiftCells].Style.Numberformat.Format = "0";

			ws1.Cells[row, CellTotalActivity + shiftCells].Formula = $"=AVERAGE({string.Join(",", averageFormulaTotalActivity)})";
			ws1.Cells[row, CellTotalActivity + shiftCells].Style.Numberformat.Format = "0";
		}

		string GroupName(string group)
		{
			if (group == "H")
			{
				return "H2O";
			}
			else if (group == "S")
			{
				return "SUC";
			}
			else if(group == "0")
			{
				return "ETOH";
			}
			else if(group == "A")
			{
				return "ALCP";
			}
			else
			{
				return group;
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

		Color GetColor(string group)
		{
			if (group == "H")
			{
				return Color.FromArgb(217, 225, 242);
			}
			else if (group == "S")
			{
				return Color.FromArgb(217, 217, 217);
			}
			else if (group == "0")
			{
				return Color.FromArgb(226, 239, 218);
			}
			else if (group == "A")
			{
				return Color.FromArgb(255, 204, 255);
			}

			if (!string.IsNullOrEmpty(currentGroup) && group != currentGroup)
			{
				if (excelColor == Color.LightBlue)
				{
					excelColor = Color.LightYellow;
				}
				else if (excelColor == Color.LightYellow)
				{
					excelColor = Color.LightPink;
				}
				else if (excelColor == Color.LightPink)
				{
					excelColor = Color.LightCoral;
				}
				else if (excelColor == Color.LightCoral)
				{
					excelColor = Color.LightSalmon;
				}
				else if (excelColor == Color.LightCoral)
				{
					excelColor = Color.LightBlue;
				}
			}

			currentGroup = group;

			return excelColor;
		}

		FileInfo GetFileInfo(string file, bool deleteIfExists = true)
		{
			var fi = new FileInfo("d:\\" + file);
			if (deleteIfExists && fi.Exists)
			{
				fi.Delete();  // ensures we create a new workbook
			}
			return fi;
		}
	}
}
