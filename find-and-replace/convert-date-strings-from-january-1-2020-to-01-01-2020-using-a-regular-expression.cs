using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with date strings in the format "January 1, 2020".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Meeting dates:");
        builder.Writeln("January 1, 2020");
        builder.Writeln("February 15, 2021");
        builder.Writeln("March 3, 2022");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that captures month name, day and year.
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})\b",
                                    RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback to format the date.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateReplacer();

        // Perform the replacement. The replacement string argument is ignored because the callback sets it.
        int replacedCount = loaded.Range.Replace(dateRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No date strings were replaced.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that converts a matched date string to "MM/dd/yyyy".
    private class DateReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, int> MonthMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "January",   1 }, { "February",  2 }, { "March",     3 },
            { "April",     4 }, { "May",       5 }, { "June",      6 },
            { "July",      7 }, { "August",    8 }, { "September", 9 },
            { "October",  10 }, { "November", 11 }, { "December", 12 }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract captured groups: month name, day, year.
            string monthName = args.Match.Groups[1].Value;
            string dayStr = args.Match.Groups[2].Value;
            string yearStr = args.Match.Groups[3].Value;

            if (!MonthMap.TryGetValue(monthName, out int monthNumber))
                return ReplaceAction.Skip; // Should not happen.

            int day = int.Parse(dayStr);
            // Build the replacement string in MM/dd/yyyy format.
            args.Replacement = $"{monthNumber:D2}/{day:D2}/{yearStr}";
            return ReplaceAction.Replace;
        }
    }
}
