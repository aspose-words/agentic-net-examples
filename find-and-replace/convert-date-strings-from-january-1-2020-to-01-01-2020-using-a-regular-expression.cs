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

        // Define a regex that matches the full month name, day and year.
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})\b",
                                    RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback that formats the date.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateReplacer();

        // Perform the replacement. The replacement string is ignored when a callback is used.
        int replacedCount = doc.Range.Replace(dateRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No date strings were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: write the resulting text to the console for verification.
        Console.WriteLine("Replacements performed: " + replacedCount);
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());
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
            // args.Match.Value contains something like "January 1, 2020".
            string[] parts = args.Match.Value.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 3)
                return ReplaceAction.Skip; // Unexpected format.

            string monthName = parts[0];
            string day = parts[1];
            string year = parts[2];

            if (!MonthMap.TryGetValue(monthName, out string monthNumber))
                return ReplaceAction.Skip; // Unknown month.

            // Ensure day is two digits.
            if (int.TryParse(day, out int dayInt))
                day = dayInt.ToString("D2");
            else
                return ReplaceAction.Skip;

            string formatted = $"{monthNumber}/{day}/{year}";
            args.Replacement = formatted;
            return ReplaceAction.Replace;
        }
    }
}
