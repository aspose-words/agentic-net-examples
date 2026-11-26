using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with known text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document containing the word old.");
        // Save the document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options for case‑sensitive matching.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("old", "new", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one case‑sensitive replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
