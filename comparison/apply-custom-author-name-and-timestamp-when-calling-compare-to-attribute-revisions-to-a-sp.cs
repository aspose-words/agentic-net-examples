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

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited paragraph with changes.");

        // Define custom author name and timestamp for the comparison.
        string customAuthor = "CustomUser";
        DateTime customDate = new DateTime(2023, 1, 1, 12, 0, 0);

        // Perform the comparison. Revisions will be attributed to the custom author and date.
        original.Compare(revised, customAuthor, customDate);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Output revision details to the console.
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision by '{rev.Author}' on {rev.DateTime:u}: {rev.RevisionType}");
        }

        // Save the document that now contains the revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(outputPath);
    }
}
