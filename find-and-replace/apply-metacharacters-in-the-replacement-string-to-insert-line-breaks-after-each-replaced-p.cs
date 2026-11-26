using System;
using Aspose.Words;
using Aspose.Words.Replacing;   // Required for FindReplaceOptions and related types

public class Program
{
    public static void Main()
    {
        // Create a sample document with paragraphs that contain the target word.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with target.");
        builder.Writeln("Second paragraph with target.");
        builder.Writeln("Third paragraph without the keyword.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Text to find and replacement string that includes a paragraph break metacharacter (\r).
        const string findText = "target";
        const string replaceText = "replaced\r";

        // Perform the find-and-replace operation using FindReplaceOptions.
        var options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace(findText, replaceText, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
