using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add sample text containing the word "color" in different cases.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("She loves the Colour of autumn leaves.");
        builder.Writeln("A rainbow displays many colors.");
        builder.Writeln("No mention of the word here.");

        // Configure find-and-replace options for case‑insensitive matching.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // Ignore case when searching.
        };

        // Perform the replacement: change every occurrence of "color" to "colour".
        int replacementCount = doc.Range.Replace("color", "colour", options);

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No occurrences of 'color' were found to replace.");

        // Define an output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docx");

        // Save the modified document.
        doc.Save(outputPath);

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
