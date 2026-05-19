using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document containing the word "color" in different cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("She likes the Colour of roses.");
        builder.Writeln("COLOR is often used in design.");
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Set up find/replace options for a case‑insensitive operation.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // Ignore case when searching.
        };

        // Perform the replacement: "color" → "colour".
        int replacedCount = loaded.Range.Replace("color", "colour", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
