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
        // Create a sample document with various date range formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Event from 01/02/2023 to 03/04/2023.");
        builder.Writeln("Meeting: 2023-01-02 - 2023-03-04.");
        builder.Writeln("Period: Jan 2, 2023 - Mar 4, 2023.");
        builder.Writeln("Dates: 2 Jan 2023 – 4 Mar 2023.");

        // Regular expression that captures two dates separated by common delimiters.
        Regex dateRangeRegex = new Regex(
            @"(?<date1>\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|[A-Za-z]{3,9}\s\d{1,2},\s\d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})\b)\s*(?:to|[-–])\s*(?<date2>\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|[A-Za-z]{3,9}\s\d{1,2},\s\d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})\b)",
            RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback that formats the dates.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateRangeStandardizer();

        // Perform the replacement. The callback will set the proper replacement string.
        int replacedCount = doc.Range.Replace(dateRangeRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No date ranges were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine($"Replaced {replacedCount} date range(s). Output saved to '{outputPath}'.");
    }

    // Callback that parses the two captured dates and rewrites them in a unified format.
    private class DateRangeStandardizer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string date1Str = args.Match.Groups["date1"].Value;
            string date2Str = args.Match.Groups["date2"].Value;

            if (TryParseDate(date1Str, out DateTime d1) && TryParseDate(date2Str, out DateTime d2))
            {
                // Unified format: yyyy-MM-dd - yyyy-MM-dd
                args.Replacement = $"{d1:yyyy-MM-dd} - {d2:yyyy-MM-dd}";
            }
            else
            {
                // If parsing fails, keep the original text.
                args.Replacement = args.Match.Value;
            }

            return ReplaceAction.Replace;
        }

        private static bool TryParseDate(string input, out DateTime date)
        {
            // Try several common date formats.
            string[] formats = {
                "M/d/yyyy", "MM/dd/yyyy", "M-d-yyyy", "MM-dd-yyyy",
                "yyyy-M-d", "yyyy-MM-dd", "yyyy/M/d", "yyyy/MM/dd",
                "MMM d, yyyy", "MMMM d, yyyy", "d MMM yyyy", "d MMMM yyyy"
            };

            return DateTime.TryParseExact(
                input,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out date);
        }
    }
}
