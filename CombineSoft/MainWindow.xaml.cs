using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

				foreach (var item in SelectedFiles)
				{
					var file = File.ReadAllLines(item);
					allFiles.Add(new FileData(file));
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
	}
}