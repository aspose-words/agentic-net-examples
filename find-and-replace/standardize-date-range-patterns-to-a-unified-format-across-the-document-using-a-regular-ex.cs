using System;
using System.Collections.Generic;
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
        builder.Writeln("Meeting dates: Jan 1 - Feb 5, 2021.");
        builder.Writeln("Another range: 01/01/2021 to 02/05/2021.");
        builder.Writeln("Third range: 2021-01-01 – 2021-02-05.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regex that matches the three date‑range patterns used above.
        string rangePattern = @"(?:(?:\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b\s*\d{1,2}\s*[-–]\s*\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b\s*\d{1,2},\s*\d{4})|(?:\d{2}/\d{2}/\d{4}\s*(?:to|[-–])\s*\d{2}/\d{2}/\d{4})|(?:\d{4}-\d{2}-\d{2}\s*[–-]\s*\d{4}-\d{2}-\d{2}))";

        // Set up the replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateRangeReplacer()
        };

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = loaded.Range.Replace(new Regex(rangePattern, RegexOptions.IgnoreCase), string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No date ranges were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that converts any matched date range to the format "yyyy-MM-dd to yyyy-MM-dd".
    private class DateRangeReplacer : IReplacingCallback
    {
        // Regex that extracts individual dates from a matched range.
        private static readonly Regex DateExtractor = new Regex(
            @"\b(?:\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},?\s*\d{4})\b",
            RegexOptions.IgnoreCase);

        // Accepted date formats for parsing.
        private static readonly string[] DateFormats = new[]
        {
            "yyyy-MM-dd",
            "MM/dd/yyyy",
            "MMM d, yyyy",
            "MMM d,yyyy",
            "MMM dd, yyyy",
            "MMM dd,yyyy"
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the two dates inside the matched range.
            MatchCollection dateMatches = DateExtractor.Matches(args.Match.Value);
            if (dateMatches.Count != 2)
                return ReplaceAction.Skip; // Unexpected format; leave unchanged.

            // Parse both dates.
            if (!TryParseDate(dateMatches[0].Value, out DateTime start) ||
                !TryParseDate(dateMatches[1].Value, out DateTime end))
                return ReplaceAction.Skip; // Parsing failed; leave unchanged.

            // Build the unified replacement string.
            args.Replacement = $"{start:yyyy-MM-dd} to {end:yyyy-MM-dd}";
            return ReplaceAction.Replace;
        }

        private static bool TryParseDate(string text, out DateTime date)
        {
            // Remove any trailing commas that may appear after month‑day.
            string cleaned = text.Trim().TrimEnd(',');
            return DateTime.TryParseExact(cleaned, DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out date);
        }
    }
}
