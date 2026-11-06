using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing various case variations of the word "color".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("Color is important in design.");
        builder.Writeln("COLOR can be bright.");
        builder.Writeln("A colorful world.");

        // Save the sample document to a local file.
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options for a case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // Ignore case when matching.
        };

        // Perform the replacement: "color" → "colour".
        int replacedCount = loaded.Range.Replace("color", "colour", options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
