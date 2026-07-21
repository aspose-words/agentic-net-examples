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

        // Normal text that will be replaced.
        builder.Writeln("Hello world!");

        // Insert a field whose field code contains the word "Hello".
        // This field will be ignored during replacement when IgnoreFields is true.
        builder.InsertField("MERGEFIELD", "Hello");

        // Save the source document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find/replace to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace "Hello" with "Hi". Only the normal text should be replaced.
        int replacedCount = loaded.Range.Replace("Hello", "Hi", options);

        // Verify that exactly one replacement occurred (the field was ignored).
        if (replacedCount != 1)
            throw new InvalidOperationException($"Expected 1 replacement, but got {replacedCount}.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loaded.Save(outputPath);
    }
}
