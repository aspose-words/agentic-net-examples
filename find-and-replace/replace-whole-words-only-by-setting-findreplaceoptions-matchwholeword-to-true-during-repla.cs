using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The cat sat on the catalog.");
        builder.Writeln("A catapult is not a cat.");
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Configure find/replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true // Ensures only standalone words are replaced.
        };

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace("cat", "dog", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole-word replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loadedDoc.Save(outputPath);
    }
}
