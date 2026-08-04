using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindAndReplaceExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with text containing the word "color" in different cases.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("The color of the sky is blue.");
            builder.Writeln("Color is an important visual attribute.");
            builder.Writeln("A colorful world is vibrant.");

            // Save the sample document to a local file.
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document we just saved.
            Document loadedDoc = new Document(inputPath);

            // Configure find‑replace options for a case‑insensitive search.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false // ignore case when searching for the pattern
            };

            // Perform the replacement: change all occurrences of "color" to "colour".
            int replacementCount = loadedDoc.Range.Replace("color", "colour", options);

            // Ensure that at least one replacement was made.
            if (replacementCount == 0)
                throw new InvalidOperationException("Expected at least one replacement, but none were performed.");

            // Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);

            // Output the result count (optional, non‑interactive).
            Console.WriteLine($"Replacements made: {replacementCount}");
        }
    }
}
