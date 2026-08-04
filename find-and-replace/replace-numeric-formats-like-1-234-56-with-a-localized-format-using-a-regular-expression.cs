using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with numbers in US format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sales figures are 1,234.56, 78,910.12 and 3,000.");
        builder.Writeln("Another line with 12,345.67 for testing.");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loadedDoc = new Document(inputPath);

        // Regular expression to match numbers with optional thousand separators and a decimal part.
        Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Set up find-and-replace options with a custom callback that localizes the number format.
        FindReplaceOptions options = new FindReplaceOptions(new NumberLocalizer());

        // Perform the replacement. The replacement string is ignored because the callback provides the value.
        int replacedCount = loadedDoc.Range.Replace(numberRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }

    // Callback that converts a US‑style number (e.g., 1,234.56) to a German‑style format (e.g., 1.234,56).
    private class NumberLocalizer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Parse the matched number using invariant culture (comma as thousand separator, dot as decimal separator).
            if (double.TryParse(args.Match.Value,
                                NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture,
                                out double number))
            {
                // Format the number using the target culture (German in this example).
                CultureInfo targetCulture = new CultureInfo("de-DE");
                string localized = number.ToString("N", targetCulture);
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
