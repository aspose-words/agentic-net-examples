using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    // Callback that standardizes date range matches.
    private class DateRangeStandardizer : IReplacingCallback
    {
        // Stores information about each replacement for reporting.
        public List<ReplacementInfo> Replacements { get; } = new List<ReplacementInfo>();

        // Mapping of month names to numeric values.
        private static readonly Dictionary<string, int> MonthMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "January", 1 }, { "Jan", 1 },
            { "February", 2 }, { "Feb", 2 },
            { "March", 3 }, { "Mar", 3 },
            { "April", 4 }, { "Apr", 4 },
            { "May", 5 },
            { "June", 6 }, { "Jun", 6 },
            { "July", 7 }, { "Jul", 7 },
            { "August", 8 }, { "Aug", 8 },
            { "September", 9 }, { "Sep", 9 }, { "Sept", 9 },
            { "October", 10 }, { "Oct", 10 },
            { "November", 11 }, { "Nov", 11 },
            { "December", 12 }, { "Dec", 12 }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the regex match.
            Match match = args.Match;
            string month1Str = match.Groups["Month1"].Value;
            string day1Str = match.Groups["Day1"].Value;
            string month2Str = match.Groups["Month2"].Success ? match.Groups["Month2"].Value : month1Str;
            string day2Str = match.Groups["Day2"].Value;
            string yearStr = match.Groups["Year"].Value;

            // Convert to integers.
            int month1 = MonthMap[month1Str];
            int day1 = int.Parse(day1Str);
            int month2 = MonthMap[month2Str];
            int day2 = int.Parse(day2Str);
            int year = int.Parse(yearStr);

            // Build standardized format: yyyy-MM-dd to yyyy-MM-dd
            string replacement = $"{year:D4}-{month1:D2}-{day1:D2} to {year:D4}-{month2:D2}-{day2:D2}";

            // Record replacement info.
            Replacements.Add(new ReplacementInfo
            {
                OriginalText = match.Value,
                StandardizedText = replacement,
                Position = args.MatchNode.GetText().IndexOf(match.Value, StringComparison.Ordinal)
            });

            // Perform the replacement.
            args.Replacement = replacement;
            return ReplaceAction.Replace;
        }
    }

    // Simple DTO for JSON reporting.
    private class ReplacementInfo
    {
        public string OriginalText { get; set; } = string.Empty;
        public string StandardizedText { get; set; } = string.Empty;
        public int Position { get; set; }
    }

    public static void Main()
    {
        // Step 1: Create a sample document with various date range formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Project timeline:");
        builder.Writeln("Jan 5 - Jan 12, 2023");
        builder.Writeln("Feb 20 - 25, 2023");
        builder.Writeln("Mar 1 – Mar 3, 2023"); // Note: en dash
        builder.Writeln("April 15 - April 20, 2023");
        builder.Writeln("May 31 - Jun 5, 2023");
        doc.Save("input.docx");

        // Step 2: Load the document for processing.
        Document loaded = new Document("input.docx");

        // Regular expression to match date ranges.
        // Supports formats like:
        //   Jan 5 - Jan 12, 2023
        //   Feb 20 - 25, 2023
        //   Mar 1 – Mar 3, 2023 (en dash)
        //   April 15 - April 20, 2023
        //   May 31 - Jun 5, 2023
        string pattern = @"\b(?<Month1>[A-Za-z]+)\s+(?<Day1>\d{1,2})\s*[-–]\s*(?:(?<Month2>[A-Za-z]+)\s*)?(?<Day2>\d{1,2}),?\s*(?<Year>\d{4})\b";
        Regex regex = new Regex(pattern, RegexOptions.Compiled);

        // Step 3: Set up replacement callback.
        DateRangeStandardizer callback = new DateRangeStandardizer();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback,
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Step 4: Perform the replacement.
        int replacedCount = loaded.Range.Replace(regex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date range replacement.");

        // Step 5: Save the modified document.
        loaded.Save("output.docx");

        // Step 6: Write a JSON report of the replacements.
        string jsonReport = JsonConvert.SerializeObject(callback.Replacements, Formatting.Indented);
        File.WriteAllText("replacements-report.json", jsonReport);

        // Validation: ensure report file was created.
        if (!File.Exists("replacements-report.json"))
            throw new InvalidOperationException("Replacement report was not created.");
    }
}
