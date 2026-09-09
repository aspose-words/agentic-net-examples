using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        // Insert a field with code QUOTE and result "Hello again!".
        builder.InsertField("QUOTE", "Hello again!");

        // Save the document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find-and-replace.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Replace "Hello" with "Greetings".
        int replaced = loaded.Range.Replace("Hello", "Greetings", options);
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replacements made: {replaced}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
