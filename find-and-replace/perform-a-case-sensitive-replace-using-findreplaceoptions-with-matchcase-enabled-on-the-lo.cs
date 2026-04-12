using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample_Modified.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with mixed‑case text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a Test.");
        builder.Writeln("hello world! Another line.");

        // Save the document so we can load it later.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // -----------------------------------------------------------------
        // 3. Set up case‑sensitive find‑and‑replace options.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true // Enable case sensitivity.
        };

        // Replace only the exact case "Hello" with "Hi".
        int replacementCount = loadedDoc.Range.Replace("Hello", "Hi", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No case‑sensitive matches were found to replace.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(modifiedPath);

        // Optional: output the result count (no user interaction required).
        Console.WriteLine($"Replacements performed: {replacementCount}");
    }
}
