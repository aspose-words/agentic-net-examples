using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and write sample text containing the word to replace.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello old value.");

        // Save the source document (ensures the example is self‑contained).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the saved file.
        Document loaded = new Document(inputPath);

        // Perform the find‑and‑replace operation and store the number of replacements.
        int replacementCount = loaded.Range.Replace("old", "new", new FindReplaceOptions());

        // Verify that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output the replacement count.
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
    }
}
