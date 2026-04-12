using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing dates in the format "January 1, 2020".
        builder.Writeln("The conference starts on January 1, 2020 and ends on February 12, 2021.");
        builder.Writeln("Another meeting is scheduled for March 5, 2022.");

        // Regular expression that captures month name, day and year.
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})\b",
                                    RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateReplacer();

        // Perform the replacement. The callback will supply the formatted replacement string.
        int replacements = doc.Range.Replace(dateRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacements == 0)
            throw new InvalidOperationException("No date strings were replaced.");

        // Save the modified document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple verification that the expected format exists in the document text.
        string text = doc.GetText();
        if (!text.Contains("01/01/2020") || !text.Contains("02/12/2021") || !text.Contains("03/05/2022"))
            throw new InvalidOperationException("Date conversion did not produce the expected results.");
    }

    // Callback that converts a matched date string to "MM/dd/yyyy".
    private class DateReplacer : IReplacingCallback
    {
        // Mapping from month name to month number.
        private static readonly Dictionary<string, int> MonthMap = new()
        {
            { "January",   1 }, { "February",  2 }, { "March",     3 },
            { "April",     4 }, { "May",       5 }, { "June",      6 },
            { "July",      7 }, { "August",    8 }, { "September", 9 },
            { "October",  10 }, { "November", 11 }, { "December", 12 }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract captured groups: month, day, year.
            string monthName = args.Match.Groups[1].Value;
            string dayStr = args.Match.Groups[2].Value;
            string yearStr = args.Match.Groups[3].Value;

            // Convert month name to its numeric representation.
            if (!MonthMap.TryGetValue(monthName, out int monthNumber))
                monthNumber = 0; // Fallback, should not happen with the regex.

            // Build the replacement string in MM/dd/yyyy format.
            string replacement = $"{monthNumber:D2}/{int.Parse(dayStr):D2}/{yearStr}";
            args.Replacement = replacement;

            return ReplaceAction.Replace;
        }
    }
}
