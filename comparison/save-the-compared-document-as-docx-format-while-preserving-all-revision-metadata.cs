using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");                     // Line that will be changed.
        builderOriginal.Writeln("This line will stay the same.");   // Unchanged line.

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");             // Modified line.
        builderRevised.Writeln("This line will stay the same.");   // Same as original.
        builderRevised.Writeln("An extra line added.");            // New line.

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that at least one revision was generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Save the compared document as DOCX, preserving all revision metadata.
        original.Save("ComparedDocument.docx");
    }
}
