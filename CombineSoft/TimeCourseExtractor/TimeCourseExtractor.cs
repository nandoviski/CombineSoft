using System.Collections.Generic;
using System.Linq;

namespace CombineSoft
{
	public class TimeCourseExtractor
	{
		public string FilePath { get; }
		public static int MultiplierCount = 24;
		public static int TimeInSeconds = 300;
		public readonly List<TimeCount> TimeCountPerAction;
		public string ErrorMessage { get; private set; }
		public bool HasError => !string.IsNullOrEmpty(ErrorMessage);
		public string Subject { get; private set; }

		public TimeCourseExtractor(string filePath, string[] file)
		{
			FilePath = filePath;
			TimeCountPerAction = new List<TimeCount>();
			Populate(file);
		}

		public class TimeCount
		{
			public string Action { get; set; }
			public Dictionary<int, double> Times;
			double expectedTotal;

			public bool IsTotalMatch => CalculateTotal() == expectedTotal;

			public TimeCount(string action, double expectedTotal)
			{
				Action = action;
				this.expectedTotal = expectedTotal;
				Times = new Dictionary<int, double>();
				for (int i = 1; i <= TimeCourseExtractor.MultiplierCount; i++)
				{
					Times.Add(i * TimeCourseExtractor.TimeInSeconds, 0);
				}
			}

			public double CalculateTotal()
			{
				var result = 0.0;
				foreach (var item in Times)
				{
					result += item.Value;
				}
				return result;
			}
		}

		void Populate(string[] file)
		{
			try
			{
				var nextAction = false;
				var isFirstLine = false;
				var multiplier = 1;
				TimeCount currentTimeCount = null;
				double ETotal = 0, FTotal = 0, GTotal = 0, HTotal = 0;

				foreach (var line in file)
				{
					if (line.StartsWith("Subject:", System.StringComparison.InvariantCultureIgnoreCase))
					{
						var subLines = line.Split(':');
						Subject = subLines[1].Trim();
					}
					else if (line.StartsWith("A:", System.StringComparison.InvariantCultureIgnoreCase))
					{
						var subLines = line.Split(':');
						ETotal = double.Parse(subLines[1].TrimEnd().TrimStart());
					}
					else if (line.StartsWith("B:", System.StringComparison.InvariantCultureIgnoreCase))
					{
						var subLines = line.Split(':');
						FTotal = double.Parse(subLines[1].TrimEnd().TrimStart());
					}
					else if (line.StartsWith("C:", System.StringComparison.InvariantCultureIgnoreCase))
					{
						var subLines = line.Split(':');
						GTotal = double.Parse(subLines[1].TrimEnd().TrimStart());
					}
					else if (line.StartsWith("D:", System.StringComparison.InvariantCultureIgnoreCase))
					{
						var subLines = line.Split(':');
						HTotal = double.Parse(subLines[1].TrimEnd().TrimStart());
					}

					if (currentTimeCount != null)
					{
						var lineData = line.Split(' ');
						foreach (var item in lineData)
						{
							var currentTimeLimit = multiplier * TimeInSeconds;

							if (!string.IsNullOrEmpty(item) && !item.Contains(":") && double.TryParse(item, out var seconds))
							{
								if (seconds > 0)
								{
									while (multiplier <= MultiplierCount)
									{
										if (seconds <= currentTimeLimit)
										{
											currentTimeCount.Times[currentTimeLimit]++;
											break;
										}
										else
										{
											multiplier++;
											currentTimeLimit = multiplier * TimeInSeconds;
										}
									}
								}

								if (!isFirstLine && seconds == 0)
								{
									nextAction = true;
									break;
								}
							}
						}

						if (nextAction == true && currentTimeCount.Action == "H")
						{
							break;
						}
						isFirstLine = false;
					}

					if (line.StartsWith("E:"))
					{
						nextAction = false;
						multiplier = 1;
						TimeCountPerAction.Add(new TimeCount("E", ETotal));
						currentTimeCount = TimeCountPerAction.LastOrDefault();
						isFirstLine = true;
					}
					else if (line.StartsWith("F:"))
					{
						nextAction = false;
						multiplier = 1;
						TimeCountPerAction.Add(new TimeCount("F", FTotal));
						currentTimeCount = TimeCountPerAction.LastOrDefault();
						isFirstLine = true;
					}
					else if (line.StartsWith("G:"))
					{
						nextAction = false;
						multiplier = 1;
						TimeCountPerAction.Add(new TimeCount("G", GTotal));
						currentTimeCount = TimeCountPerAction.LastOrDefault();
						isFirstLine = true;
					}
					else if (line.StartsWith("H:"))
					{
						nextAction = false;
						multiplier = 1;
						TimeCountPerAction.Add(new TimeCount("H", HTotal));
						currentTimeCount = TimeCountPerAction.LastOrDefault();
						isFirstLine = true;
					}
				}

				foreach (var item in TimeCountPerAction)
				{
					if (!item.IsTotalMatch)
					{
						ErrorMessage = $"{FilePath}\n\nTotal of column the {item.Action} don't match with the expected value";
					}
				}
			}
			catch (System.Exception ex)
			{
				ErrorMessage = FilePath + "\n" + ex.Message;
			}
		}
	}
}
