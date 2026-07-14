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
        // Set the culture to French (France) to demonstrate localized formatting.
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");

        // Create a sample document with numbers in US format (e.g., 1,234.56).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sales figures are as follows:");
        builder.Writeln("January: 1,234.56");
        builder.Writeln("February: 7,890.12");
        builder.Writeln("March: 3,456.78");
        // Save the original document (optional, just for reference).
        doc.Save("Input.docx");

        // Define a regular expression that matches numbers with optional thousands separators and a decimal part.
        // Example matches: 1,234.56 , 7,890 , 123.45 , 12,345,678.90
        Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Configure find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumberLocalizer();

        // Perform the replacement. The callback will supply the localized replacement string.
        int replacedCount = doc.Range.Replace(numberRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that converts the matched US-formatted number to the current culture's format.
    private class NumberLocalizer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Parse the matched number using invariant culture (comma as thousands separator, dot as decimal separator).
            if (double.TryParse(args.Match.Value,
                                NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture,
                                out double number))
            {
                // Format the number using the current culture with two decimal places.
                string localized = number.ToString("N2", CultureInfo.CurrentCulture);
                args.Replacement = localized;
            }
            else
            {
                // If parsing fails, keep the original text.
                args.Replacement = args.Match.Value;
            }

            return ReplaceAction.Replace;
        }
    }
}
