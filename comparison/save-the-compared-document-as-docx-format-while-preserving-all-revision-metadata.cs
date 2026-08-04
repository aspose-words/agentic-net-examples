using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a deliberate change.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Perform the comparison. The original document will now contain revision metadata.
        original.Compare(revised, "John Doe", DateTime.Now);

        // Verify that at least one revision was generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");
        }

        // Define the output path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");

        // Save the compared document preserving all revision information.
        original.Save(outputPath);
    }
}
