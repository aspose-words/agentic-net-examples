using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the category rules
using Newtonsoft.Json; // Required by the category rules

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file with known content.
        string inputPath = "input.docx";
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is the old value that will be replaced. Another old occurrence.");
        sampleDoc.Save(inputPath);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Perform a literal string replacement.
        string findText = "old";
        string replaceText = "new";
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(findText, replaceText, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");
        }

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: indicate success (no interactive input required).
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
