using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class DateRangeStandardizer : IReplacingCallback
{
    private readonly List<string> _matches = new List<string>();
    public IReadOnlyList<string> Matches => _matches;

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Record the original matched text.
        _matches.Add(args.Match.Value);

        // Extract the two date parts using named groups.
        string date1Text = args.Match.Groups["date1"].Value;
        string date2Text = args.Match.Groups["date2"].Value;

        // Try to parse both dates using a set of known formats.
        if (!TryParseDate(date1Text, out DateTime date1) ||
            !TryParseDate(date2Text, out DateTime date2))
        {
            // If parsing fails, keep the original text.
            args.Replacement = args.Match.Value;
            return ReplaceAction.Skip;
        }

        // Build the unified format: yyyy-MM-dd to yyyy-MM-dd
        string replacement = $"{date1:yyyy-MM-dd} to {date2:yyyy-MM-dd}";
        args.Replacement = replacement;
        return ReplaceAction.Replace;
    }

    private bool TryParseDate(string text, out DateTime date)
    {
        // Define possible date formats.
        string[] formats = new[]
        {
            "yyyy-MM-dd",
            "MM/dd/yyyy",
            "MMMM d, yyyy",
            "MMM d, yyyy"
        };

        // Try exact parsing first.
        if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture,
                                   DateTimeStyles.None, out date))
        {
            return true;
        }

        // Fallback to general parsing.
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with various date range formats.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("The project runs from January 1, 2020 - February 5, 2020.");
        builder.Writeln("Another range: 01/01/2020 to 02/05/2020.");
        builder.Writeln("Legacy format: 2020-01-01 – 2020-02-05.");
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Regular expression to match two dates separated by dash, en‑dash, or the word "to".
        Regex dateRangeRegex = new Regex(
            @"(?<date1>\b(?:\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[/-]\d{2}[/-]\d{4}|[A-Za-z]+ \d{1,2}, \d{4})\b)\s*(?:-|to|–|—)\s*(?<date2>\b(?:\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[/-]\d{2}[/-]\d{4}|[A-Za-z]+ \d{1,2}, \d{4})\b)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Set up the replacement callback.
        DateRangeStandardizer callback = new DateRangeStandardizer();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the regex replacement. The replacement string is ignored because the callback supplies it.
        int replacedCount = doc.Range.Replace(dateRangeRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date range replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
