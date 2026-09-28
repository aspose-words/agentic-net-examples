using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package, not used directly in this example.

public class Program
{
    public static void Main()
    {
        // Create a sample document containing the word "color" in different cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("She likes the Color red.");
        builder.Writeln("COLOR is often used in design.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options for a case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // Ignore case when searching.
        };

        // Perform the replacement of "color" with "colour".
        int replacedCount = loaded.Range.Replace("color", "colour", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
