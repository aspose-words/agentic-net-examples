using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and insert a MERGEFIELD.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with a field:");
        builder.InsertField("MERGEFIELD MyField \\* MERGEFORMAT");
        builder.Writeln(); // Add a line break.

        // Save the initial document (optional, for inspection).
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Capture the field before replacement.
        Field? fieldBefore = doc.Range.Fields.Count > 0 ? doc.Range.Fields[0] : null;
        if (fieldBefore == null)
            throw new InvalidOperationException("Expected a field in the document.");

        string codeBefore = fieldBefore.GetFieldCode();
        string resultBefore = fieldBefore.Result;

        // Set up FindReplaceOptions to include fields (default behavior).
        FindReplaceOptions options = new FindReplaceOptions
        {
            // IgnoreFields = false is the default; explicitly set for clarity.
            IgnoreFields = false
        };

        // Replace "MyField" with "YourField" only in field codes.
        int replacedCount = doc.Range.Replace("MyField", "YourField", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in field codes.");

        // Capture the field after replacement.
        Field? fieldAfter = doc.Range.Fields.Count > 0 ? doc.Range.Fields[0] : null;
        if (fieldAfter == null)
            throw new InvalidOperationException("Field disappeared after replacement.");

        string codeAfter = fieldAfter.GetFieldCode();
        string resultAfter = fieldAfter.Result;

        // Output verification information.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Field code before: {codeBefore}");
        Console.WriteLine($"Field result before: {resultBefore}");
        Console.WriteLine($"Field code after: {codeAfter}");
        Console.WriteLine($"Field result after (should be unchanged): {resultAfter}");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }
}
