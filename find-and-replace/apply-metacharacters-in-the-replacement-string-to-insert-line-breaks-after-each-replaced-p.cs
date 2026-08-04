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

        // -------------------------------------------------
        // Create a sample document with several paragraphs.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph contains the token ReplaceMe.");
        builder.Writeln("Second paragraph also contains ReplaceMe.");
        builder.Writeln("Third paragraph does not have the token.");

        // Save the source document.
        doc.Save(inputPath);

        // -------------------------------------------------
        // Load the document and perform find-and-replace.
        // The replacement string uses the meta‑character &p to insert a paragraph break.
        // -------------------------------------------------
        Document loaded = new Document(inputPath);
        int replacedCount = loaded.Range.Replace("ReplaceMe", "ReplaceMe&p", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
