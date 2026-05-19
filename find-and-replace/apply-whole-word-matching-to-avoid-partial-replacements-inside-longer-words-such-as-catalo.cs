using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Retained as required by the task, though not used in this example.

public class Program
{
    public static void Main()
    {
        // Create a sample document with a target word and a longer word that should not be altered.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The library catalogue is updated.");
        builder.Writeln("The catalogue123 entry should not be changed.");

        // Save the sample document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find‑and‑replace processing.
        Document loaded = new Document(inputPath);

        // Configure the replace operation to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the whole word "catalogue" with "catalog".
        int replaced = loaded.Range.Replace("catalogue", "catalog", options);

        // Verify that at least one replacement occurred.
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
