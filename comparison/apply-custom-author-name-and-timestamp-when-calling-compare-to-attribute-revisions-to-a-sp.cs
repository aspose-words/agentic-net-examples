using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It contains some sample text.");

        // Create the revised document with modifications.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited paragraph."); // Changed text.
        builderRevised.Writeln("It contains some sample text."); // Same as original.
        builderRevised.Writeln("An additional line was added.");   // New line.

        // Define custom author name and timestamp for the comparison.
        string customAuthor = "CustomUser";
        DateTime customDate = new DateTime(2023, 12, 31, 23, 59, 59, DateTimeKind.Utc);

        // Perform the comparison. Revisions will be attributed to the custom author and timestamp.
        original.Compare(revised, customAuthor, customDate);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Output revision details to the console.
        Console.WriteLine($"Total revisions: {original.Revisions.Count}");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime:u}");
            Console.WriteLine($"  Affected text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(outputPath);
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
