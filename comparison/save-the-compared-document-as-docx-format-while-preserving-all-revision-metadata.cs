using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");
        builderOriginal.Writeln("This line will stay the same.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words!");
        builderRevised.Writeln("This line will stay the same.");

        // Perform the comparison. The original document will receive revision metadata.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Save the compared document (which contains all revisions) as DOCX.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(outputPath, SaveFormat.Docx);

        // Optional: write the revision count to the console for verification.
        Console.WriteLine($"Comparison complete. Revisions count: {original.Revisions.Count}");
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
