using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with normal text and a field that contains the word "Hello".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");                     // Normal text.
        builder.InsertField("MERGEFIELD", "Hello");          // Field containing the same word.

        // Save the document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find/replace to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace "Hello" with "Hi". The occurrence inside the field will be ignored.
        int replacedCount = loaded.Range.Replace("Hello", "Hi", options);

        // Ensure that at least one replacement occurred outside the field.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement outside fields.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output simple verification information.
        Console.WriteLine($"Replacements performed (ignoring fields): {replacedCount}");
        Console.WriteLine($"Output document saved to: {Path.GetFullPath(outputPath)}");
    }
}
