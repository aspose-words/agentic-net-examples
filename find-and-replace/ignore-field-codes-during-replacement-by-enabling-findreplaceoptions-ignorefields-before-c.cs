using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with normal text and a field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        builder.InsertField("QUOTE", "Hello again!");

        // Save the document to a local file.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find/replace to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace "Hello" with "Greetings" while ignoring fields.
        int replacedCount = loaded.Range.Replace("Hello", "Greetings", options);

        // Expect exactly one replacement (the one outside the field).
        if (replacedCount != 1)
            throw new InvalidOperationException($"Expected 1 replacement, but got {replacedCount}.");

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");
        loaded.Save(outputPath);
    }
}
