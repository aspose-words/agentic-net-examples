using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");
        builder1.Writeln("This line will stay the same.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello Aspose.Words world!"); // Modified line.
        builder2.Writeln("This line will stay the same."); // Unchanged line.
        builder2.Writeln("Additional line added."); // New line.

        // Define custom author name and timestamp for the revisions.
        string customAuthor = "John Doe";
        DateTime customDate = new DateTime(2023, 1, 1, 12, 0, 0);

        // Perform the comparison, attributing revisions to the custom author and timestamp.
        original.Compare(revised, customAuthor, customDate);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Output details of each revision.
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime:u}");
            Console.WriteLine($"Type: {rev.RevisionType}");
            Console.WriteLine($"Text: {rev.ParentNode.GetText().Trim()}");
            Console.WriteLine();
        }

        // Save the document that now contains the revisions.
        const string outputFile = "ComparedDocument.docx";
        original.Save(outputFile);
        Console.WriteLine($"Compared document saved to '{outputFile}'.");
    }
}
