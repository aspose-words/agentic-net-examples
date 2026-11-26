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
        builderOriginal.Writeln("It contains several lines of text.");
        builderOriginal.Writeln("The quick brown fox jumps over the lazy dog.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original document."); // Same line.
        builderRevised.Writeln("It contains several lines of modified text."); // Modified line.
        builderRevised.Writeln("The quick brown fox jumps over the lazy dog."); // Same line.
        builderRevised.Writeln("An additional line has been added."); // New line.

        // Perform comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the compared document preserving all revision metadata.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");
        original.Save(outputPath, SaveFormat.Docx);
    }
}
