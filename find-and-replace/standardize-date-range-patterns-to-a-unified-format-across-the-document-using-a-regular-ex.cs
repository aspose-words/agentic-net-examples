using System;
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

        builder.Writeln("Report period: 01/15/2021 - 02/20/2021");
        builder.Writeln("Fiscal year: March 5, 2020 – April 10, 2020");
        builder.Writeln("Project timeline: 2021-03-01 to 2021-06-30");
        builder.Writeln("Another range: 12-31-2020 - 01-15-2021");

        // Save the original document (optional, for inspection).
        const string inputPath = "DateRangesInput.docx";
        doc.Save(inputPath);

        // Define a regex that captures two dates separated by a dash or the word "to".
        // This pattern handles:
        //   - MM/dd/yyyy or MM-dd-yyyy
        //   - dd-MM-yyyy or dd/MM/yyyy
        //   - MMMM d, yyyy (e.g., March 5, 2020)
        //   - yyyy-MM-dd
        //   - The separator can be '-', '–' (en dash), or the word "to".
        string pattern = @"(?<date1>(?:\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})|(?:[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})|(?:\d{4}[\/-]\d{1,2}[\/-]\d{1,2}))\s*(?:-|–|to)\s*(?<date2>(?:\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})|(?:[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})|(?:\d{4}[\/-]\d{1,2}[\/-]\d{1,2}))";

        Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

        // Set up find/replace options with a custom callback to format dates.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateRangeReplacer();

        // Perform the replacement. The replacement string is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(regex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No date ranges were found for replacement.");

        // Save the modified document.
        const string outputPath = "DateRangesStandardized.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Replaced {replacedCount} date range(s).");
        Console.WriteLine($"Input document:  {Path.GetFullPath(inputPath)}");
        Console.WriteLine($"Output document: {Path.GetFullPath(outputPath)}");
    }

    // Callback that converts each matched date range to the format "yyyy-MM-dd to yyyy-MM-dd".
    private class DateRangeReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the captured date strings.
            string dateStr1 = args.Match.Groups["date1"].Value.Trim();
            string dateStr2 = args.Match.Groups["date2"].Value.Trim();

            // Parse the dates using DateTime.Parse, which handles the formats defined above.
            if (!DateTime.TryParse(dateStr1, out DateTime date1) ||
                !DateTime.TryParse(dateStr2, out DateTime date2))
            {
                // If parsing fails, keep the original text.
                args.Replacement = args.Match.Value;
                return ReplaceAction.Replace;
            }

            // Build the unified replacement string.
            args.Replacement = $"{date1:yyyy-MM-dd} to {date2:yyyy-MM-dd}";
            return ReplaceAction.Replace;
        }
    }
}
