using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It has two paragraphs.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the edited document."); // Modified first line.
        builderRevised.Writeln("It has three paragraphs.");    // Modified second line.
        builderRevised.Writeln("Additional paragraph added."); // New paragraph.

        // Compare the documents, generating revisions in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were actually created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were generated after comparison.");
        }

        // Save the compared document in the legacy DOC format, preserving all revision metadata.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.doc");
        original.Save(outputPath, SaveFormat.Doc);

        // Inform the user about the result.
        Console.WriteLine($"Comparison complete. Revisions count: {original.Revisions.Count}");
        Console.WriteLine($"Document saved with revisions to: {outputPath}");
    }
}
