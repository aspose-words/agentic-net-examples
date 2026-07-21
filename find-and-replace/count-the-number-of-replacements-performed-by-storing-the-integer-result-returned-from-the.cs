using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Paths for the input and output documents
        const string inputPath = "sample.docx";
        const string outputPath = "sample_replaced.docx";

        // Create a sample document if it does not already exist
        if (!File.Exists(inputPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document. Replace the word 'sample' with 'example'.");
            builder.Writeln("Another sample line with the word sample multiple times: sample, sample.");
            doc.Save(inputPath);
        }

        // Load the document from disk
        Document loadedDoc = new Document(inputPath);

        // Perform find-and-replace and capture the number of replacements
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementCount = loadedDoc.Range.Replace("sample", "example", options);

        // Ensure that at least one replacement was made
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document
        loadedDoc.Save(outputPath);

        // Report the number of replacements performed
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
    }
}
