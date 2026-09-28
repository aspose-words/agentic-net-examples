using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ComparisonSaveDocExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a few lines of text.");
        builderOriginal.Writeln("The quick brown fox jumps over the lazy dog.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original document."); // Same line.
        builderRevised.Writeln("It contains a few lines of modified text."); // Modified line.
        builderRevised.Writeln("The quick brown fox jumps over the lazy dog."); // Same line.
        builderRevised.Writeln("An extra line has been added."); // New line.

        // Perform comparison. Provide author name and current date/time.
        original.Compare(revised, "ComparisonAuthor", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the compared document in DOC format while preserving revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.doc");
        original.Save(outputPath, SaveFormat.Doc);

        // Optional: write a simple console message indicating success.
        Console.WriteLine($"Comparison complete. Document saved as DOC with {original.Revisions.Count} revision(s) at:");
        Console.WriteLine(outputPath);
    }
}
