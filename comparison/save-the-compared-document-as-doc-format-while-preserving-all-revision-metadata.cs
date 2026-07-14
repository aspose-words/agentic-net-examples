using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This line will stay the same.");

        // Create the revised document with a deliberate change.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello brave new world."); // Modified text.
        builderRevised.Writeln("This line will stay the same."); // Unchanged text.

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that at least one revision was generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the compared document as a binary DOC file, preserving all revision metadata.
        const string outputFile = "ComparedDocument.doc";
        original.Save(outputFile, SaveFormat.Doc);

        // Output basic information (no interactive prompts).
        Console.WriteLine($"Revisions count: {original.Revisions.Count}");
        Console.WriteLine($"Document saved to: {outputFile}");
    }
}
