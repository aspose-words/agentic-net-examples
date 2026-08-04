using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FindReplaceIgnoreFieldsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add regular text that will be replaced.
        builder.Writeln("Hello world!");

        // Insert a field whose result also contains the word "Hello".
        // The field code is QUOTE and the result text is "Hello again!".
        builder.InsertField("QUOTE", "Hello again!");

        // Save the original document for reference.
        const string originalPath = "original.docx";
        doc.Save(originalPath);

        // Set up find‑replace options to ignore text inside fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // When true, the replace operation will skip any text that is part of a field
            // (both the field code and the field result). This demonstrates that the
            // field result "Hello again!" is not altered.
            IgnoreFields = true
        };

        // Perform the replacement: change "Hello" to "Greetings".
        int replacedCount = doc.Range.Replace("Hello", "Greetings", options);

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Update fields so that any changes to field codes would be reflected in the result.
        // In this case the field code was not changed because we ignored fields,
        // but the call demonstrates the typical workflow.
        doc.UpdateFields();

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output simple verification to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Original document saved as: {Path.GetFullPath(originalPath)}");
        Console.WriteLine($"Modified document saved as: {Path.GetFullPath(outputPath)}");
        Console.WriteLine("Document text after replacement:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
