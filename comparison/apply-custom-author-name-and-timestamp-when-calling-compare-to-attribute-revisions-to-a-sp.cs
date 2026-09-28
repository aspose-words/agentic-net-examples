using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It will be compared against the revised version.");

        // Create the revised document with modifications.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised paragraph."); // changed text
        builderRevised.Writeln("It will be compared against the original version."); // changed text

        // Define custom author name and timestamp for the comparison.
        string customAuthor = "John Doe";
        DateTime customDate = new DateTime(2023, 1, 1, 12, 0, 0);

        // Perform the comparison. Revisions will be attributed to the custom author and date.
        original.Compare(revised, customAuthor, customDate);

        // Verify that revisions have the expected author and date.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Total revisions detected: {revisionCount}");

        foreach (Revision revision in original.Revisions)
        {
            // Output revision details.
            Console.WriteLine($"Revision Type: {revision.RevisionType}");
            Console.WriteLine($"Author: {revision.Author}");
            Console.WriteLine($"Date: {revision.DateTime}");

            // Get the text associated with the revision via its parent node.
            string revisionText = revision.ParentNode?.GetText() ?? string.Empty;
            Console.WriteLine($"Text: {revisionText}");
            Console.WriteLine(new string('-', 40));
        }

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
