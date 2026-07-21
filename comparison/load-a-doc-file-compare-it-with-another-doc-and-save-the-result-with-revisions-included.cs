using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonOutput");
        Directory.CreateDirectory(outputDir);

        string originalPath = Path.Combine(outputDir, "original.docx");
        string revisedPath = Path.Combine(outputDir, "revised.docx");
        string resultPath = Path.Combine(outputDir, "comparison-result.docx");

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This is the original document.");
        original.Save(originalPath);

        // Create the revised document with differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello world!"); // Slight change.
        builderRevised.Writeln("This is the revised document with changes."); // Modified line.
        revised.Save(revisedPath);

        // Load the documents (optional, they are already in memory).
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // Perform comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (docOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result which includes the revisions.
        docOriginal.Save(resultPath);
    }
}
