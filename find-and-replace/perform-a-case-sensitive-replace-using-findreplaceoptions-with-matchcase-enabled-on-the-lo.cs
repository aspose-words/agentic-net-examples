using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // Create a sample document with text that includes different capitalizations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple is red. apple is green. APPLE is tasty.");
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loadedDoc = new Document(inputPath);

        // Configure find‑replace options for case‑sensitive matching.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true
        };

        // Perform the replacement: only the exact case "Apple" will be replaced.
        int replacedCount = loadedDoc.Range.Replace("Apple", "Orange", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}
