using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceColorWithColour
{
    static void Main()
    {
        // Create a new document with sample text containing the word "color" in various cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("Color is a primary attribute.");
        builder.Writeln("A colorful world.");

        // Configure find/replace options for case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // ignore case when matching "color"
        };

        // Perform the replacement throughout the whole document.
        doc.Range.Replace("color", "colour", options);

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
