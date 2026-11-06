using System;
using System.IO;
using Aspose.Words;

public class ComparisonDemo
{
    public static void Main()
    {
        // Prepare a folder for the sample documents.
        string basePath = Directory.GetCurrentDirectory();
        string demoFolder = Path.Combine(basePath, "ComparisonDemo");
        Directory.CreateDirectory(demoFolder);

        // Paths for the original and revised documents.
        string originalPath = Path.Combine(demoFolder, "Original.docx");
        string revisedPath = Path.Combine(demoFolder, "Revised.docx");
        string resultPath = Path.Combine(demoFolder, "ComparisonResult.docx");

        // Create the original document with some content.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("Hello world.");
        originalDoc.Save(originalPath);

        // Create the revised document with a difference.
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.Writeln("Hello revised world.");
        revisedDoc.Save(revisedPath);

        // Load the documents from the file system.
        Document docToCompare = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // Perform the comparison. Revisions will be added to docToCompare.
        docToCompare.Compare(docRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (docToCompare.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the document that now contains the revisions.
        docToCompare.Save(resultPath);

        // Output a simple summary.
        Console.WriteLine($"Comparison completed. Revisions found: {docToCompare.Revisions.Count}");
        Console.WriteLine($"Result document saved to: {resultPath}");
    }
}
