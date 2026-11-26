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
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with different content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Perform the comparison so that revisions are placed in the revised document.
        // Revisions are added to the document on which Compare is called.
        revised.Compare(original, "Comparer", DateTime.Now);

        // Verify that revisions exist in the revised document.
        if (revised.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions in the revised document, but none were found.");
        }

        // Save both documents to the current directory.
        string outputDir = Directory.GetCurrentDirectory();
        original.Save(Path.Combine(outputDir, "original.docx"));
        revised.Save(Path.Combine(outputDir, "revised_with_revisions.docx"));
    }
}
