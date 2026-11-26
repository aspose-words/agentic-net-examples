using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder originalBuilder = new DocumentBuilder(original);
        originalBuilder.Writeln("This is the original text.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder revisedBuilder = new DocumentBuilder(revised);
        revisedBuilder.Writeln("This is the revised text.");

        // Compare the documents; revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions, which removes the revision markup.
        original.AcceptAllRevisions();

        // Verify that the revisions collection is now empty.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Revisions collection should be empty after accepting.");
        }

        // Save the final document without any revision marks.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "final.docx");
        original.Save(outputPath);
    }
}
