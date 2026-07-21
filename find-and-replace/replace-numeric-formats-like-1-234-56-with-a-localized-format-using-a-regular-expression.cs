using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json; // Included as a required package

namespace FindAndReplaceNumericLocalization
{
    // Callback that swaps thousand and decimal separators (e.g., "1,234.56" → "1.234,56")
    public class NumberLocalizer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Original matched number (e.g., "1,234.56")
            string original = args.Match.Value;

            // Swap separators: ',' → temporary '#', '.' → ',', '#' → '.'
            string swapped = original.Replace(",", "#")
                                    .Replace(".", ",")
                                    .Replace("#", ".");

            args.Replacement = swapped;
            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document with numbers in US format.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("The sales figures are as follows:");
            builder.Writeln("Revenue: 1,234.56 USD");
            builder.Writeln("Cost: 987,654.32 USD");
            builder.Writeln("Profit: 12,345.67 USD");

            // Save the original document (optional, for inspection).
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Define a regex that matches numbers with optional thousand separators and a decimal part.
            // Example matches: 1,234.56 , 987,654.32 , 12,345.67
            Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

            // Set up find/replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new NumberLocalizer()
            };

            // Perform the replacement.
            int replacedCount = doc.Range.Replace(numberRegex, string.Empty, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one numeric replacement.");

            // Save the modified document.
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            // Simple verification output (optional).
            Console.WriteLine($"Replaced {replacedCount} numeric occurrences.");
            Console.WriteLine($"Input document saved to: {Path.GetFullPath(inputPath)}");
            Console.WriteLine($"Output document saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
