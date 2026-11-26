using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing the word "color" in different cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("Color is an important visual attribute.");
        builder.Writeln("A colorful world is vibrant.");

        // Configure find‑replace options for a case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // Ignore case when searching.
        };

        // Perform the replacement: "color" → "colour".
        int replacedCount = doc.Range.Replace("color", "colour", options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none were performed.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
