using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Collections.Generic;

public class DateRangeReplacer : IReplacingCallback
{
    private static readonly string[] DateFormats = new[]
    {
        "MM/dd/yyyy",
        "yyyy-MM-dd",
        "MMM d, yyyy",
        "MMM dd, yyyy"
    };

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The match contains two named groups: date1 and date2.
        var match = args.Match;
        string date1Str = match.Groups["date1"].Value;
        string date2Str = match.Groups["date2"].Value;

        if (DateTime.TryParseExact(date1Str, DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime date1) &&
            DateTime.TryParseExact(date2Str, DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime date2))
        {
            // Standardize to ISO format with "to" separator.
            args.Replacement = $"{date1:yyyy-MM-dd} to {date2:yyyy-MM-dd}";
        }
        else
        {
            // If parsing fails, keep the original text.
            args.Replacement = match.Value;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with various date range formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The project runs from Jan 1, 2020 - Feb 5, 2020.");
        builder.Writeln("Another period: 01/01/2020 to 02/05/2020.");
        builder.Writeln("Legacy format: 2020-01-01 – 2020-02-05.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regex to match date ranges with different separators and date formats.
        Regex dateRangeRegex = new Regex(
            @"(?<date1>\b(?:\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}|[A-Za-z]{3}\s+\d{1,2},\s*\d{4})\b)\s*(?:-|to|–)\s*(?<date2>\b(?:\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}|[A-Za-z]{3}\s+\d{1,2},\s*\d{4})\b)",
            RegexOptions.Compiled);

        // Set up find-replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateRangeReplacer()
        };

        // Perform the replacement. The replacement string is ignored when a callback is used.
        int replacedCount = loaded.Range.Replace(dateRangeRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date range replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
