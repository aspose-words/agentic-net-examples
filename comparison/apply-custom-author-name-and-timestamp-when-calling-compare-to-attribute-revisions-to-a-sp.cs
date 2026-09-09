using System;
using Aspose.Words;

public class CompareWithCustomAuthor
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited paragraph with a change.");

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
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}");
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime:u}");
            Console.WriteLine($"Changed Text: \"{rev.ParentNode.GetText().Trim()}\"");
            Console.WriteLine();
        }

        // Save the compared document (contains revisions) to the local folder.
        string outputPath = "Compared.docx";
        original.Save(outputPath);
        Console.WriteLine($"Comparison document saved to: {outputPath}");
    }
}
