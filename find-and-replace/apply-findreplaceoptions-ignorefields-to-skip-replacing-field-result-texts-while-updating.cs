using System;
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
        // Insert a QUOTE field whose result contains the word "Hello".
        builder.InsertField("QUOTE", "Hello again!");

        // Save the document so we can load it later (simulates an existing file).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Configure find-and-replace to ignore text inside fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace the word "Hello" with "Greetings". The field result will be left unchanged.
        int replacementCount = loadedDoc.Range.Replace("Hello", "Greetings", options);

        // Ensure that at least one replacement occurred (outside the field).
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement outside of fields.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Optional: output a simple confirmation.
        Console.WriteLine($"Replacements made: {replacementCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
