using System;
using System.IO;
using Aspose.Words;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a line that will be changed.");

        // Create the revised document with modified content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original document.");
        builderRevised.Writeln("It contains a line that has been changed."); // Difference

        // Define custom author name and timestamp for the comparison.
        string customAuthor = "John Doe";
        DateTime customDate = new DateTime(2023, 5, 15, 10, 30, 0, DateTimeKind.Utc);

        // Perform the comparison. Revisions will be attributed to the custom author and date.
        original.Compare(revised, customAuthor, customDate);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Output revision details to the console.
        foreach (Revision revision in original.Revisions)
        {
            Console.WriteLine($"Revision Type: {revision.RevisionType}");
            Console.WriteLine($"Author: {revision.Author}");
            Console.WriteLine($"Date: {revision.DateTime:u}");

            // Retrieve the text associated with the revision, if possible.
            string revisionText = revision.ParentNode?.GetText() ?? string.Empty;
            Console.WriteLine($"Text: {revisionText.TrimEnd()}");

            Console.WriteLine(new string('-', 40));
        }

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");
        original.Save(outputPath);
    }
}
