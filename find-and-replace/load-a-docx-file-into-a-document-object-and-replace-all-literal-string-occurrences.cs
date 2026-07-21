using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file with text that contains the target string.
        var sampleDoc = new Document();
        var builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello old value.");
        builder.Writeln("Another old line.");
        const string inputFile = "input.docx";
        sampleDoc.Save(inputFile);

        // Load the document from the file system.
        var doc = new Document(inputFile);

        // Replace all occurrences of the literal string "old" with "new".
        int replacementCount = doc.Range.Replace("old", "new", new FindReplaceOptions());

        // Verify that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputFile = "output.docx";
        doc.Save(outputFile);
    }
}
