using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string basePath = Directory.GetCurrentDirectory();
        string originalPath = Path.Combine(basePath, "Original.docx");
        string revisedPath = Path.Combine(basePath, "Revised.docx");
        string resultPath = Path.Combine(basePath, "ComparedWithRevisions.docx");

        // -----------------------------------------------------------------
        // Create the original document with some content.
        // -----------------------------------------------------------------
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        // Save so we can load it later as a file.
        original.Save(originalPath);

        // -----------------------------------------------------------------
        // Create the revised document with a deliberate change.
        // -----------------------------------------------------------------
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");
        revised.Save(revisedPath);

        // -----------------------------------------------------------------
        // Load the two documents from disk (simulating external files).
        // -----------------------------------------------------------------
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // -----------------------------------------------------------------
        // Perform the comparison. Revisions will be added to docOriginal.
        // -----------------------------------------------------------------
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (docOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");
        }

        // -----------------------------------------------------------------
        // Save the comparison result which contains the tracked changes.
        // -----------------------------------------------------------------
        docOriginal.Save(resultPath);

        // Optional: write a simple console message (no user interaction required).
        Console.WriteLine($"Comparison complete. Revisions count: {docOriginal.Revisions.Count}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
