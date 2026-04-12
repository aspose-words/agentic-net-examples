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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing date ranges in different formats.
        builder.Writeln("Date ranges to be standardized:");
        builder.Writeln("01/01/2023 to 01/05/2023");
        builder.Writeln("Jan 1 - Jan 5, 2023");
        builder.Writeln("2023-01-01 – 2023-01-05"); // en‑dash between dates

        // Prepare find‑replace options with a custom callback that formats the matched range.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateRangeStandardizer()
        };

        // Pattern 1: numeric dates separated by "to" or a dash.
        Regex numericPattern = new Regex(
            @"(?<date1>\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*(?:to|[-–])\s*(?<date2>\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
            RegexOptions.IgnoreCase);

        // Pattern 2: month name dates like "Jan 1 - Jan 5, 2023".
        Regex textualPattern = new Regex(
            @"(?<date1>\w+\s+\d{1,2})\s*(?:-|–|to)\s*(?<date2>\w+\s+\d{1,2}),\s*(?<year>\d{4})",
            RegexOptions.IgnoreCase);

        // Perform replacements.
        int countNumeric = doc.Range.Replace(numericPattern, string.Empty, options);
        int countTextual = doc.Range.Replace(textualPattern, string.Empty, options);
        int totalReplacements = countNumeric + countTextual;

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No date ranges were found to replace.");

        // Ensure the output folder exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StandardizedDates.docx");
        doc.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replaced {totalReplacements} date range(s). Output saved to: {outputPath}");
    }

    // Callback that parses the matched date range and rewrites it in ISO format.
    private class DateRangeStandardizer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            Match match = args.Match;

            // Try numeric pattern first.
            if (match.Groups["date1"].Success && match.Groups["date2"].Success && !match.Groups["year"].Success)
            {
                if (TryParseDate(match.Groups["date1"].Value, out DateTime start) &&
                    TryParseDate(match.Groups["date2"].Value, out DateTime end))
                {
                    args.Replacement = $"{start:yyyy-MM-dd} to {end:yyyy-MM-dd}";
                    return ReplaceAction.Replace;
                }
            }

            // Try textual pattern with a shared year.
            if (match.Groups["date1"].Success && match.Groups["date2"].Success && match.Groups["year"].Success)
            {
                string year = match.Groups["year"].Value;
                string date1Str = $"{match.Groups["date1"].Value} {year}";
                string date2Str = $"{match.Groups["date2"].Value} {year}";

                if (TryParseDate(date1Str, out DateTime start) &&
                    TryParseDate(date2Str, out DateTime end))
                {
                    args.Replacement = $"{start:yyyy-MM-dd} to {end:yyyy-MM-dd}";
                    return ReplaceAction.Replace;
                }
            }

            // If parsing fails, skip replacement.
            return ReplaceAction.Skip;
        }

        private static bool TryParseDate(string input, out DateTime date)
        {
            // Accept several common date formats.
            string[] formats = {
                "M/d/yyyy", "MM/dd/yyyy", "M-d-yyyy", "MM-dd-yyyy",
                "yyyy-M-d", "yyyy-MM-dd", "d/M/yyyy", "dd/MM/yyyy",
                "MMM d yyyy", "MMMM d yyyy", "MMM d, yyyy", "MMMM d, yyyy"
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
