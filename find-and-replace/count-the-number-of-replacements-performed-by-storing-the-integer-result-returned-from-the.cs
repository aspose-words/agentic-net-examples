using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that will be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is quick.");
        // Save the sample document to a local file.
        const string inputPath = "sample_input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loadedDoc = new Document(inputPath);

        // Perform a find-and-replace operation and capture the number of replacements.
        const string findText = "quick";
        const string replaceText = "swift";
        int replacementCount = loadedDoc.Range.Replace(findText, replaceText, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "sample_output.docx";
        loadedDoc.Save(outputPath);

        // Output the count to the console (optional, not required for validation).
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
    }
}
