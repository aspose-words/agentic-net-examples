using System;
using Aspose.Words;

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

        // Perform the comparison so that revisions appear in the revised (new) document.
        // By calling Compare on the revised document and passing the original as the source,
        // the revisions are stored in the revised document.
        revised.Compare(original, "Comparer", DateTime.Now);

        // Verify that revisions are present in the revised document.
        int revisionCount = revised.Revisions.Count;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException(
                "Expected revisions in the revised document, but none were found.");
        }

        // Save both documents.
        string outputDir = Environment.CurrentDirectory;
        original.Save(System.IO.Path.Combine(outputDir, "Original.docx"));
        revised.Save(System.IO.Path.Combine(outputDir, "Revised_With_Revisions.docx"));

        // Output revision count to console (non‑interactive).
        Console.WriteLine($"Revisions in revised document: {revisionCount}");
    }
}
