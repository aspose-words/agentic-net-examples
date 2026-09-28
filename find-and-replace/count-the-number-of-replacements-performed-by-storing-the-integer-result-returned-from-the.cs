using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FindAndReplaceCountExample
{
    public static void Main()
    {
        // Create a sample document with text that will be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the old value that will be replaced. Another old occurrence.");
        doc.Save("input.docx");

        // Load the document from the file.
        Document loaded = new Document("input.docx");

        // Perform the replacement and capture the number of replacements made.
        int replacedCount = loaded.Range.Replace("old", "new", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Output the replacement count.
        Console.WriteLine($"Number of replacements performed: {replacedCount}");
    }
}
