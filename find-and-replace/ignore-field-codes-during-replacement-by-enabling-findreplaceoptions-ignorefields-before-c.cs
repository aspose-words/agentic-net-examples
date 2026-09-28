using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with normal text and a field containing the target word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a PLACEHOLDER in normal text.");
        // Insert a MERGEFIELD whose field code includes the word PLACEHOLDER.
        builder.InsertField("MERGEFIELD PLACEHOLDER \\* MERGEFORMAT");
        builder.Writeln(); // Add a line break.

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find-and-replace.
        Document loaded = new Document(inputPath);

        // Configure find-replace options to ignore field codes.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that at least one replacement occurred (the normal text only).
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
