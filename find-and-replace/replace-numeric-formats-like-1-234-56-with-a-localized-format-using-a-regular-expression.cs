using System;
using System.Globalization;
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

        // Add sample text containing numbers in US format (e.g., 1,234.56).
        builder.Writeln("The total cost is 1,234.56 dollars.");
        builder.Writeln("Another value: 12,345.00 and 987.65.");
        builder.Writeln("No change for 1234 or 1.234,56.");

        // Define a regular expression that matches numbers with optional thousands separators and a decimal part.
        // Example matches: 1,234.56 , 12,345.00 , 987.65
        Regex numberPattern = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new NumberLocalizer(new CultureInfo("de-DE")) // German format: 1.234,56
        };

        // Perform the replacement.
        int replacementCount = doc.Range.Replace(numberPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No numeric values were replaced.");

        // Save the modified document.
        const string outputPath = "LocalizedNumbers.docx";
        doc.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacementCount} numeric occurrence(s). Document saved to '{outputPath}'.");
    }

    // Callback that converts each matched US‑style number to the target culture's format.
    private class NumberLocalizer : IReplacingCallback
    {
        private readonly CultureInfo _targetCulture;

        public NumberLocalizer(CultureInfo targetCulture)
        {
            _targetCulture = targetCulture ?? throw new ArgumentNullException(nameof(targetCulture));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Parse the matched number using invariant culture (comma as thousands separator, dot as decimal separator).
            if (double.TryParse(args.Match.Value,
                                NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture,
                                out double value))
            {
                // Format the number using the target culture.
                args.Replacement = value.ToString(_targetCulture);
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
