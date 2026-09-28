using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        const string inputFile = "input.docx";
        const string outputFile = "output.docx";

        // Create a sample DOCX file with text that contains the target string.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is the old value.");
        builder.Writeln("Another line with old text to replace.");
        sampleDoc.Save(inputFile);

        // Load the created document.
        Document doc = new Document(inputFile);

        // Replace all literal occurrences of "old" with "new".
        int replacedCount = doc.Range.Replace("old", "new", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        doc.Save(outputFile);
    }
}
