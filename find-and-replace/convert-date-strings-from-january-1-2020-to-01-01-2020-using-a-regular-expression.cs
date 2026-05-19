using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with dates in the format "January 1, 2020".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The first meeting is on January 1, 2020.");
        builder.Writeln("The second meeting is on February 12, 2021.");
        builder.Writeln("Another date: March 5, 2022.");
        // Save the source document (required by the rules).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loadedDoc = new Document(inputPath);

        // Regular expression to match full month name, day and year.
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})\b",
                                    RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback that builds the new date string.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateReplacer()
        };

        // Perform the replacement. The replacement string is supplied by the callback, so we pass an empty string.
        int replacedCount = loadedDoc.Range.Replace(dateRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No dates were replaced. Check the input content and regex pattern.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }

    // Callback that converts a matched date string to "MM/dd/yyyy".
    private class DateReplacer : IReplacingCallback
    {
        // Mapping from month name to its numeric representation.
        private static readonly Dictionary<string, string> MonthMap = new()
        {
            { "January",   "01" }, { "February",  "02" }, { "March",     "03" },
            { "April",     "04" }, { "May",       "05" }, { "June",      "06" },
            { "July",      "07" }, { "August",    "08" }, { "September", "09" },
            { "October",   "10" }, { "November",  "11" }, { "December",  "12" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract captured groups: month name, day, year.
            string monthName = args.Match.Groups[1].Value;
            string day = args.Match.Groups[2].Value;
            string year = args.Match.Groups[3].Value;

            // Convert month name to number, pad day with leading zero if needed.
            if (!MonthMap.TryGetValue(monthName, out string monthNumber))
                monthNumber = "00";

            if (day.Length == 1)
                day = "0" + day;

            // Build the replacement string in MM/dd/yyyy format.
            args.Replacement = $"{monthNumber}/{day}/{year}";
            return ReplaceAction.Replace;
        }
    }
}
