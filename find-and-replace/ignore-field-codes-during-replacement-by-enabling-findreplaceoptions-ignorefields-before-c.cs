using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add normal text that will be replaced.
        builder.Writeln("Hello world!");

        // Insert a field whose result also contains the word "Hello".
        // The QUOTE field simply displays the provided text.
        builder.InsertField("QUOTE", "Hello inside field!");

        // Add a line break after the field for readability.
        builder.Writeln();

        // Configure find/replace to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true // Do not replace text inside fields.
        };

        // Perform the replacement.
        int replacementCount = doc.Range.Replace("Hello", "Greetings", options);

        // Validate that at least one replacement occurred (the normal paragraph).
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Verify that the field's result was NOT changed.
        if (doc.Range.Fields.Count == 0)
            throw new InvalidOperationException("No fields were found in the document.");

        string fieldResult = doc.Range.Fields[0].Result;
        if (fieldResult.Contains("Greetings"))
            throw new InvalidOperationException("Field text was replaced despite IgnoreFields being true.");

        // Verify that the normal text was replaced.
        string documentText = doc.GetText();
        if (!documentText.Contains("Greetings world!"))
            throw new InvalidOperationException("Normal text was not replaced as expected.");

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Replacement completed. Output saved to: {outputPath}");
    }
}
