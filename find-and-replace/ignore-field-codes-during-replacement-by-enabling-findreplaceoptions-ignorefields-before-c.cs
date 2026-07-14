using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add normal text that will be replaced.
        builder.Writeln("Hello world!");

        // Insert a field whose field code contains the same text.
        // The field code will be ignored during replacement when IgnoreFields is true.
        builder.InsertField("QUOTE", "Hello again!");

        // Save the original document (optional, for inspection).
        doc.Save("original.docx");

        // Configure find/replace options to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Perform the replacement. Only the normal text should be changed.
        int replacedCount = doc.Range.Replace("Hello", "Greetings", options);

        // Validate that exactly one replacement occurred (the normal paragraph).
        if (replacedCount != 1)
            throw new InvalidOperationException($"Expected 1 replacement, but got {replacedCount}.");

        // Save the modified document.
        doc.Save("output.docx");

        // Output a simple confirmation.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine("Modified document saved as 'output.docx'.");
    }
}
