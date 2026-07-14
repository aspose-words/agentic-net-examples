using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with date ranges in different formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The project runs from 01/01/2020 - 02/05/2020.");
        builder.Writeln("Another period: 2020-03-10 to 2020-04-15.");
        builder.Writeln("Mixed format: 12-31-2020 to 01/15/2021.");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Regular expression that matches two dates separated by '-' or 'to'.
        // Supports MM/dd/yyyy, MM-dd-yyyy, yyyy-MM-dd and yyyy/MM/dd.
        Regex dateRangeRegex = new Regex(
            @"(?<d1>\d{2}[\/-]\d{2}[\/-]\d{4})\s*(?:-|to)\s*(?<d2>\d{2}[\/-]\d{2}[\/-]\d{4})" +
            @"|(?<d3>\d{4}[\/-]\d{2}[\/-]\d{2})\s*(?:-|to)\s*(?<d4>\d{4}[\/-]\d{2}[\/-]\d{2})",
            RegexOptions.Compiled);

        // Set up find‑replace options with a custom callback that formats the dates.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateRangeReplacer()
        };

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = loaded.Range.Replace(dateRangeRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date range replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that parses the two dates found by the regex and rewrites them
    // in a unified ISO format: yyyy-MM-dd to yyyy-MM-dd.
    private class DateRangeReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            Match match = args.Match;

            // Determine which capture groups were matched.
            string first = match.Groups["d1"].Success ? match.Groups["d1"].Value : match.Groups["d3"].Value;
            string second = match.Groups["d2"].Success ? match.Groups["d2"].Value : match.Groups["d4"].Value;

            DateTime dt1 = ParseDate(first);
            DateTime dt2 = ParseDate(second);

            args.Replacement = $"{dt1:yyyy-MM-dd} to {dt2:yyyy-MM-dd}";
            return ReplaceAction.Replace;
        }

        private static DateTime ParseDate(string text)
        {
            string[] formats = { "MM/dd/yyyy", "MM-dd-yyyy", "yyyy-MM-dd", "yyyy/MM/dd" };
            return DateTime.ParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None);
        }
    }
}
