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
        builderOriginal.Writeln("It contains two paragraphs.");

        // Create the revised document that has differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document."); // Modified first line.
        builderRevised.Writeln("It contains two paragraphs.");   // Same as original.
        builderRevised.Writeln("Additional paragraph added in revised version."); // New line.

        // Compare the documents, providing author name and comparison date.
        original.Compare(revised, "ComparisonAuthor", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");
        }

        // Save the compared document as DOCX, preserving all revision metadata.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");
        original.Save(outputPath, SaveFormat.Docx);
    }
}
