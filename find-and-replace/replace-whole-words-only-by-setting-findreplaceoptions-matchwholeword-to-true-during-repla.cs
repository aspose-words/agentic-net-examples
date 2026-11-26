using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required package, not used directly in this example.

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("foxes are clever."); // This line should NOT be changed because "fox" is not a whole word here.

        // Save the sample document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the document for replacement.
        Document loaded = new Document(inputPath);

        // Configure find-replace options to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the whole word "fox" with "cat".
        int replacedCount = loaded.Range.Replace("fox", "cat", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole-word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath, SaveFormat.Docx);
    }
}
