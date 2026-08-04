using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the category rules
using Newtonsoft.Json; // Required by the category rules

public class FindAndReplaceCountExample
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeFindReplaceDemo");
        Directory.CreateDirectory(workFolder);

        // Define file paths for the input and output documents.
        string inputPath = Path.Combine(workFolder, "input.docx");
        string outputPath = Path.Combine(workFolder, "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with text that contains the target pattern.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the old value.");
        builder.Writeln("Another line with the old value.");
        builder.Writeln("No match on this line.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Perform a find-and-replace operation and capture the replacement count.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions(); // default options
        int replacementCount = loadedDoc.Range.Replace("old", "new", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // Output the result count (no interactive prompts required).
        Console.WriteLine($"Number of replacements performed: {replacementCount}");
    }
}
