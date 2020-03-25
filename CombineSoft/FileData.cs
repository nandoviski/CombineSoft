using System;
using System.Text.RegularExpressions;

namespace CombineSoft
{
	public class FileData
	{
		public FileData(string[] file)
		{
			Populate(file);
		}

		public string FileName { get; set; }
		public DateTime StartDate { get; set; }
		public TimeSpan StartTime { get; set; }
		public DateTime EndDate { get; set; }
		public TimeSpan EndTime { get; set; }
		public string Subject { get; set; }
		public string Experiment { get; set; }
		public string Group { get; set; }
		public int Box { get; set; }
		public string Msn { get; set; }

		public double Active { get; set; }
		public double Inactive { get; set; }
		public double Infusions { get; set; }
		public double TotalActivity { get; set; }
		public double TotalTime { get; set; }
		public double Activity1 { get; set; }
		public double Activity2 { get; set; }
		public double Activity3 { get; set; }
		public double Activity4 { get; set; }

		public string Gender { get; private set; }
		public int RatNumber { get; private set; }

		public void Populate(string[] file)
		{
			string startDate = string.Empty, endDate = string.Empty, startTime = string.Empty, endTime = string.Empty;

			foreach (var item in file)
			{
				var line = item.Split(new[] { ':' }, 2);
				if (line.Length >= 2)
				{
					var line1 = line[1].TrimStart();

					switch (line[0].ToUpper())
					{
						case "FILE":
							FileName = line1;
							break;
						case "START DATE":
							startDate = line1.Trim();
							break;
						case "END DATE":
							endDate = line1.Trim();
							break;
						case "SUBJECT":
							Subject = line1;
							Gender = Regex.Replace(line1, @"[\d-]", string.Empty);
							if (int.TryParse(Regex.Replace(line1, "[^0-9.]", string.Empty), out var ratNumber))
							{
								RatNumber = ratNumber;
							}
							break;
						case "EXPERIMENT":
							Experiment = line1;
							break;
						case "GROUP":
							Group = line1;
							break;
						case "BOX":
							Box = int.Parse(line1);
							break;
						case "MSN":
							Msn = line1;
							break;
						case "START TIME":
							startTime = line1.Trim();
							break;
						case "END TIME":
							endTime = line1.Trim();
							break;
						case "A":
							Active = double.Parse(line1.Trim());
							break;
						case "B":
							Inactive = double.Parse(line1.Trim());
							break;
						case "C":
							Infusions = double.Parse(line1.Trim());
							break;
						case "D":
							TotalActivity = double.Parse(line1.Trim());
							break;
						case "N":
							TotalTime = double.Parse(line1.Trim());
							break;
						case "R":
							Activity1 = double.Parse(line1.Trim());
							break;
						case "S":
							Activity2 = double.Parse(line1.Trim());
							break;
						case "T":
							Activity3 = double.Parse(line1.Trim());
							break;
						case "U":
							Activity4 = double.Parse(line1.Trim());
							break;
					}

					if (line[0] == "Z")
					{
						break;
					}
				}
			}
			
			if (!string.IsNullOrEmpty(startDate))
			{
				var date = startDate.Split('/');
				StartDate = new DateTime(int.Parse(date[2]), int.Parse(date[0]), int.Parse(date[1]));
			}

			if (!string.IsNullOrEmpty(startTime))
			{
				var time = startTime.Split(':');
				StartTime = new TimeSpan(int.Parse(time[0]), int.Parse(time[1]), int.Parse(time[2]));
			}

			if (!string.IsNullOrEmpty(endDate))
			{
				var date = endDate.Split('/');
				EndDate = new DateTime(int.Parse(date[2]), int.Parse(date[0]), int.Parse(date[1]));
			}

			if (!string.IsNullOrEmpty(endTime))
			{
				var time = endTime.Split(':');
				EndTime = new TimeSpan(int.Parse(time[0]), int.Parse(time[1]), int.Parse(time[2]));
			}
		}
	}
}
