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

        // -----------------------------------------------------------------
        // 1. Create a sample document with several paragraphs containing a placeholder.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write three paragraphs, each containing the word "_PLACEHOLDER_".
        builder.Writeln("_PLACEHOLDER_");
        builder.Writeln("This paragraph does not contain the token.");
        builder.Writeln("_PLACEHOLDER_");
        builder.Writeln("_PLACEHOLDER_");

        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and perform find-and-replace.
        //    The replacement string uses the meta‑character "&p" to insert a paragraph break
        //    after each occurrence of the placeholder.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Replace the placeholder with itself followed by a paragraph break.
        int replacedCount = loaded.Range.Replace("_PLACEHOLDER_", "_PLACEHOLDER_&p", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
