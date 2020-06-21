using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Win32;

namespace CombineSoft
{
	public partial class MainWindow : Window
	{
		string[] SelectedFiles;

		public MainWindow()
		{
			InitializeComponent();
		}

		void Search_Click(object sender, RoutedEventArgs e)
		{
			var openFileDialog1 = new OpenFileDialog();
			openFileDialog1.Multiselect = true;

			if (openFileDialog1.ShowDialog() == true)
			{
				SelectedFiles = openFileDialog1.FileNames;
				txtSearch.Text = string.Join("\n", SelectedFiles);
			}
		}

		void Open_Click(object sender, RoutedEventArgs e)
		{
			if (SelectedFiles?.Length > 0)
			{
				var allFiles = new List<FileData>();
				var errors = new StringBuilder();
				foreach (var item in SelectedFiles)
				{
					var file = File.ReadAllLines(item);
					var fileData = new FileData(item, file);
					if (!fileData.HasError)
					{
						allFiles.Add(fileData);
					}
					else
					{
						errors.AppendLine(fileData.ErrorMessage);
					}
				}

				if (!string.IsNullOrEmpty(errors.ToString()))
				{
					MessageBox.Show("Some files couldn't be processed:\n\n" + errors.ToString(), "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				}

				this.dataGrid1.ItemsSource = allFiles;

				var result = new ExcelUtil().CreateExcel(allFiles, txtFilePath.Text);
				if (result)
				{
					var msgResult = MessageBox.Show("Excel created on '"+ txtFilePath.Text + "'\nDo you wish to open?", "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);
					if (msgResult == MessageBoxResult.Yes)
					{
						Process.Start(txtFilePath.Text);
					}
				}
			}
		}

		void Find_Click(object sender, RoutedEventArgs e)
		{
			var saveFileDialog1 = new SaveFileDialog();
			saveFileDialog1.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls|All files (*.*)|*.*";
			
			if (saveFileDialog1.ShowDialog() == true)
			{
				txtFilePath.Text = saveFileDialog1.FileName;
			}
		}

		void TimeCourseExtractor_Click(object sender, RoutedEventArgs e)
		{
			if (SelectedFiles?.Length > 0)
			{
				var allFiles = new List<TimeCourseExtractor>();
				var errors = new StringBuilder();
				foreach (var item in SelectedFiles)
				{
					var file = File.ReadAllLines(item);
					var fileData = new TimeCourseExtractor(item, file);

					if (fileData.HasError)
					{
						errors.AppendLine(fileData.ErrorMessage);
					}
					//else
					//{
						allFiles.Add(fileData);
					//}
				}

				if (!string.IsNullOrEmpty(errors.ToString()))
				{
					MessageBox.Show("Some files with issues:\n\n" + errors.ToString(), "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				}

				var result = new TimeCourseExtractorExcelUtil().CreateExcel(allFiles, txtFilePath.Text);
				if (result)
				{
					var msgResult = MessageBox.Show("Excel created on '" + txtFilePath.Text + "'\nDo you wish to open?", "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);
					if (msgResult == MessageBoxResult.Yes)
					{
						Process.Start(txtFilePath.Text);
					}
				}
			}
		}
	}
}