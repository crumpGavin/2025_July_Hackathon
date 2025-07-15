// Usings
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

// Asssociate with the ArchManu namespace
namespace gSharp.ArchManu
{
    [Transaction(TransactionMode.Manual)]
    // Command class - convert PaxRay JSON file
    public class Cmd_PaxRayConvert : IExternalCommand
    {
        // Your JSON File here
        
        // JSON Properties
        private readonly string PROPERTY_APPNAME = "ApplicationName";
        private readonly string PROPERTY_STARTTIME = "TimeStarted";
        private readonly string PROPERTY_KEYSTROKES = "KeyStrokes";
        private readonly string PROPERTY_MOUSECLICKS = "MouseClicks";

        // Tracker variables
        private DateTime TIME_LAST;
        private string APPNAME_LAST;

        // Execute our command
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Read JSON file
            var filePath = SelectJson();
            if (!filePath.EndsWith(".json")) { return Result.Failed; }
            string json = File.ReadAllText(filePath);
            var dataArray = JArray.Parse(json).ToList();

            // See if we have an ignore list
            var directoryPath = Path.GetDirectoryName(filePath);
            var ignorePath = Path.Combine(directoryPath, "IgnoreList.tsv");
            var ignoreApps = new List<string>();

            if (File.Exists(ignorePath) && FileIsAccessible(filePath))
            {
                ignoreApps = ReadFileAsList(ignorePath);
            }

            // Get first item, initialize variables
            var firstItem = dataArray[0];
            APPNAME_LAST = GetStringValue(firstItem, PROPERTY_APPNAME);
            TIME_LAST = GetDateTimeValue(firstItem, PROPERTY_STARTTIME);
            var firstKeys = GetIntValue(firstItem, PROPERTY_KEYSTROKES);
            var firstClicks = GetIntValue(firstItem, PROPERTY_MOUSECLICKS);

            // Initialize dictionaries per application
            var timeInApp = new Dictionary<string, double>();
            timeInApp[APPNAME_LAST] = 0.0;
            var keysInApp = new Dictionary<string, int>();
            keysInApp[APPNAME_LAST] = firstKeys;
            var clicksInApp = new Dictionary<string, int>();
            clicksInApp[APPNAME_LAST] = firstClicks;
            var switchKeys = new Dictionary<string, int>();

            // Event lists to build
            var eventApps = new List<string>() { APPNAME_LAST };
            var eventTimes = new List<double>();
            var eventKeys = new List<int>() { firstKeys };
            var eventClicks = new List<int>() { firstClicks };
            var eventSwitches = new List<string>();
            var eventStarts = new List<string>() { ToDateTimeString(TIME_LAST) };
            var eventEnds = new List<string>() { };
            var eventNumber = 0;

            // Remove first item now that we have stored it
            var objectCount = dataArray.Count;
            dataArray.RemoveAt(objectCount - 1);

            // For each item in the JSON array...
            foreach (var item in dataArray)
            {
                // Get the app name
                var appName = GetStringValue(item, PROPERTY_APPNAME);

                // Cancel if it's in the ignore list
                if (ignoreApps.Contains(appName)) { continue; }

                // Add to dictionaries if it doesn't exist
                if (!keysInApp.ContainsKey(appName))
                {
                    timeInApp[appName] = 0;
                    keysInApp[appName] = 0;
                    clicksInApp[appName] = 0;
                }

                // Get keys and clicks
                var keyStrokes = GetIntValue(item, PROPERTY_KEYSTROKES);
                var clicks = GetIntValue(item, PROPERTY_MOUSECLICKS);

                // Uptick the tallies
                keysInApp[appName] += keyStrokes;
                clicksInApp[appName] += clicks;

                // If the app is the same...
                if (appName == APPNAME_LAST)
                {
                    // Add to event properties
                    eventKeys[eventNumber] += keyStrokes;
                    eventClicks[eventNumber] += clicks;
                }
                // Otherwise, new app...
                else
                {
                    // Get total time spent since last start time, add to app total
                    var endTime = GetDateTimeValue(item, PROPERTY_STARTTIME);
                    var elapsedTime = (endTime - TIME_LAST).TotalSeconds;
                    var stringTime = ToDateTimeString(endTime);
                    timeInApp[APPNAME_LAST] += elapsedTime;

                    // Yield to the event lists
                    eventApps.Add(appName);
                    eventTimes.Add(elapsedTime);
                    eventKeys.Add(keyStrokes);
                    eventClicks.Add(clicks);
                    eventStarts.Add(stringTime);
                    eventEnds.Add(stringTime);

                    // Make switch key
                    var switchKey = $"{APPNAME_LAST} > {appName}";

                    // Add if needed
                    if (!switchKeys.ContainsKey(switchKey))
                    {
                        switchKeys[switchKey] = 0;
                    }

                    // Uptick the switch
                    switchKeys[switchKey] += 1;
                    eventSwitches.Add(switchKey);

                    // New last app and start time
                    APPNAME_LAST = appName;
                    TIME_LAST = endTime;
                    eventNumber++;
                }
            }

            // Construct the event log
            var eventRows = new List<string>()
            {
                "App\tDuration\tKeyStrokes\tMouseClicks\tStartTime\tEndTime\tSwithTo"
            };

            for (int i = 0; i < eventNumber; i++)
            {
                var dataRow = $"{i}\t" +
                    $"{eventApps[i]}\t" +
                    $"{eventTimes[i]}\t" +
                    $"{eventKeys[i]}\t" +
                    $"{eventClicks[i]}\t" +
                    $"{eventStarts[i]}\t" +
                    $"{eventEnds[i]}\t" +
                    $"{eventSwitches[i]}";
                eventRows.Add(dataRow);
            }

            // Construct the app log
            var appRows = new List<string>()
            {
                "App\tKeyStrokes\tMouseClicks\tDuration"
            };

            foreach (var keyedValue in timeInApp)
            {
                var appName = keyedValue.Key;
                var appKeys = keysInApp[appName].ToString();
                var appClicks = clicksInApp[appName].ToString();
                var appTime = timeInApp[appName].ToString();
                appRows.Add($"{appName}\t{appKeys}\t{appClicks}\t{appTime}");
            }

            // Construct the switch log
            var appSwitches = new List<string>()
            {
                "AppSwitch\tFrom\tTo\tCount"
            };

            var splitString = new string[] { " > " };

            foreach (var keyedValue in switchKeys)
            {
                var switchName = keyedValue.Key;
                var switchCount = keyedValue.Value.ToString();
                var switchParts = switchName.Split(splitString, StringSplitOptions.None);
                var switchFrom = switchParts.First();
                var switchTo = switchParts.Last();
                appSwitches.Add($"{switchName}\t{switchFrom}\t{switchTo}\t{switchCount}");
            }

            // Write the event log
            var eventLogPath = Path.Combine(Path.GetDirectoryName(filePath), "EventLog.tsv");
            WriteFileAsList(eventLogPath, eventRows);

            // Write the app tallies
            var appLogPath = Path.Combine(Path.GetDirectoryName(filePath), "AppLog.tsv");
            WriteFileAsList(appLogPath, appRows);

            // Write the switch tallies
            var switchLogPath = Path.Combine(Path.GetDirectoryName(filePath), "SwitchLog.tsv");
            WriteFileAsList(switchLogPath, appSwitches);

            // Return command success
            return Result.Succeeded;
        }

        /// <summary>
        /// Attempts to get a DateTime from a JSON item.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private DateTime GetDateTimeValue(JToken item, string propertyName)
        {
            try
            {
                var timeString = item[propertyName].ToString();
                return DateTime.Parse(timeString);
            }
            catch { return DateTime.MinValue; }
        }

        /// <summary>
        /// Attempts to get an Integer from a JSON item.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>

        private int GetIntValue(JToken item, string propertyName)
        {
            try
            {
                var intString = item[propertyName].ToString();
                return StringToInt(intString);
            }
            catch { return 0; }
        }

        /// <summary>
        /// Attempts to get a String from a JSON item.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private string GetStringValue(JToken item, string propertyName)
        {
            try
            {
                return item[propertyName].ToString();
            }
            catch { return null; }
        }

        /// <summary>
        /// Converts a datetime to regular format.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private string ToDateTimeString(DateTime dateTime)
        {
            return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// Converts a string to an integer.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static int StringToInt(string text)
        {
            int x = 0;

            if (Int32.TryParse(text, out x))
            {
                return x;
            }

            return 0;
        }

        /// <summary>
        /// Writes a list of strings to a file path.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public static Result WriteFileAsList(string filePath, List<string> dataList)
        {
            if (filePath is null || !FileIsAccessible(filePath))
            {
                return Result.Failed;
            }

            using (StreamWriter writer = new StreamWriter(filePath, false))
            {
                foreach (string row in dataList)
                {
                    try
                    {
                        writer.WriteLine(row);
                    }
                    catch
                    {
                        {; }
                    }
                }
            }
            return Result.Succeeded;
        }

        /// <summary>
        /// Ensures a file is accessible.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FileIsAccessible(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return true;
            }

            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Prompts the selection of a file.
        /// </summary>
        /// <returns></returns>
        public static string SelectJson()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select a file";
                dialog.Multiselect = false;
                dialog.RestoreDirectory = true;
                dialog.Filter = "Json Files (*.json)|*.json";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    if (dialog.FileNames.Length > 0)
                    {
                        return dialog.FileNames.First();
                    }
                }

                return "";
            }
        }

        /// <summary>
        /// Reads a file as a list.
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static List<string> ReadFileAsList(string filepath)
        {
            var list = new List<string>();

            try
            {
                using (var reader = new StreamReader(filepath))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (line is not null)
                        {
                            list.Add(line);
                        }
                    }
                }
            }
            catch
            {
                ;
            }
            return list;
        }
    }
}