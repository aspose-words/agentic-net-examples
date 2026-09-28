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
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a few lines of text.");

        // Create the revised document with differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document."); // Modified line.
        builderRevised.Writeln("It contains a few lines of text."); // Same line.
        builderRevised.Writeln("An additional line is added."); // New line.

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
