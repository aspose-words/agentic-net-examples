using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original document with initial content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Compare the documents; revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that at least one revision was generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison.");
        }

        // Save the compared document as DOC format, preserving all revision metadata.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.doc");
        original.Save(outputPath, SaveFormat.Doc);

        // Reload the saved DOC file to verify that revisions are still present.
        Document loaded = new Document(outputPath);
        if (loaded.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Revisions were not preserved in the DOC file.");
        }
    }
}
