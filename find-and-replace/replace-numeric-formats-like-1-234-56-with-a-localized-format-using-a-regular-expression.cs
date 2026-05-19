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
        // Create a sample document containing numbers in US format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The total amounts are 1,234.56, 7,890.12 and 3,000.");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Regular expression that matches numbers like 1,234.56 or 3,000.
        Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Set up find‑and‑replace options with a custom callback that localizes numbers.
        FindReplaceOptions options = new FindReplaceOptions(new NumberLocalizer());

        // Perform the replacement. The callback supplies the localized replacement text.
        int replacedCount = loadedDoc.Range.Replace(numberRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        Console.WriteLine($"Replaced {replacedCount} numeric values. Output saved to '{outputPath}'.");
    }

    // Callback that converts matched US‑style numbers to the current culture format.
    private class NumberLocalizer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Parse the matched value using invariant culture (comma = thousands, dot = decimal).
            if (double.TryParse(args.Match.Value,
                                NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture,
                                out double number))
            {
                // Format the number using the current thread's culture.
                args.Replacement = number.ToString("N", CultureInfo.CurrentCulture);
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
