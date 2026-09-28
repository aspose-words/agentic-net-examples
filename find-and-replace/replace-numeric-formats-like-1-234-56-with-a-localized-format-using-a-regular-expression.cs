using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that formats matched numeric strings from US to German locale.
    private class NumericFormatCallback : IReplacingCallback
    {
        private readonly CultureInfo _sourceCulture = CultureInfo.InvariantCulture; // US style (comma thousand, dot decimal)
        private readonly CultureInfo _targetCulture = new CultureInfo("de-DE");      // German style (dot thousand, comma decimal)

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Get the matched text (e.g., "1,234.56").
            string match = args.Match.Value;

            // Parse the number using the source culture.
            if (decimal.TryParse(match, NumberStyles.Number, _sourceCulture, out decimal number))
            {
                // Format the number using the target culture.
                string replacement = number.ToString(_targetCulture);
                args.Replacement = replacement;
            }
            else
            {
                // If parsing fails, keep the original text.
                args.Replacement = match;
            }

            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a sample document with numeric values in US format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Invoice total: 1,234.56");
        builder.Writeln("Tax amount: 78.90");
        builder.Writeln("Discount: 12,345.00");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches numbers with optional thousands separators and a decimal part.
        Regex numericRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Set up find-replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericFormatCallback();

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(numericRegex, "", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric format replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
