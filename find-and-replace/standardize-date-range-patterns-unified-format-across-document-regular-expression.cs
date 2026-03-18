using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Regular expression that matches several common date formats:
        //   MM/dd/yyyy or M/d/yyyy
        //   dd-MM-yyyy or d-M-yyyy
        //   yyyy.MM.dd
        //   yyyy/MM/dd
        //   yyyy-MM-dd
        // The pattern uses named groups to identify the parts of the date.
        string pattern = @"(?<month>\d{1,2})[\/\-\.\s](?<day>\d{1,2})[\/\-\.\s](?<year>\d{2,4})|(?<year2>\d{4})[\/\-\.\s](?<month2>\d{1,2})[\/\-\.\s](?<day2>\d{1,2})";

        Regex regex = new Regex(pattern);

        // Configure replace options with a callback that builds the unified format.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateStandardizerCallback();

        // Perform the replace. The replacement string is ignored because the callback supplies it.
        doc.Range.Replace(regex, string.Empty, options);

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that converts any matched date to the ISO format yyyy-MM-dd.
    private class DateStandardizerCallback : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            Match m = e.Match;
            string year, month, day;

            // Format: month/day/year (e.g., 12/31/2021)
            if (m.Groups["year"].Success)
            {
                month = m.Groups["month"].Value.PadLeft(2, '0');
                day   = m.Groups["day"].Value.PadLeft(2, '0');
                year  = NormalizeYear(m.Groups["year"].Value);
            }
            // Format: year-month-day (e.g., 2021-12-31)
            else
            {
                year  = m.Groups["year2"].Value;
                month = m.Groups["month2"].Value.PadLeft(2, '0');
                day   = m.Groups["day2"].Value.PadLeft(2, '0');
            }

            // Build the unified representation.
            e.Replacement = $"{year}-{month}-{day}";
            return ReplaceAction.Replace;
        }

        // Helper to turn two‑digit years into four‑digit years.
        private string NormalizeYear(string y)
        {
            if (y.Length == 2)
            {
                int yr = int.Parse(y);
                // Assume 00‑49 => 2000‑2049, 50‑99 => 1950‑1999
                return yr <= 49 ? $"20{y}" : $"19{y}";
            }
            return y;
        }
    }
}
