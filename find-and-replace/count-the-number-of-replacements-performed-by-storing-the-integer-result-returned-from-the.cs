using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly in this example
using Newtonsoft.Json; // Required package, not used directly in this example

public class Program
{
    public static void Main()
    {
        // Define file names for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document with text that will be replaced.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is swift.");
        // Save the document so it can be loaded later.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document from the file system.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // Perform a find-and-replace operation.
        // The Replace method returns the number of replacements performed.
        // -----------------------------------------------------------------
        int replacementCount = loaded.Range.Replace(
            "quick",
            "swift",
            new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");
        }

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // Output the count to the console (optional, does not require user interaction).
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
    }
}
