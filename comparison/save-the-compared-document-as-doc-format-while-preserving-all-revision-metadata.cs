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
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Perform the comparison. The original document will contain revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the compared document in the legacy DOC format, preserving revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.doc");
        original.Save(outputPath, SaveFormat.Doc);
    }
}
