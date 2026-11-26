using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the first sample document (original).
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");
        builder1.Writeln("This is the original document.");

        // Save the original document to a local file.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        original.Save(originalPath, SaveFormat.Docx);

        // Create the second sample document (revised) with intentional differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world."); // Modified line.
        builder2.Writeln("This is the original document."); // Same line.
        builder2.Writeln("An extra line added in revised version."); // New line.

        // Save the revised document to a local file.
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        revised.Save(revisedPath, SaveFormat.Docx);

        // Load the two documents from the files.
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // Compare the documents. The result will be stored in docOriginal with revisions.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

        // Ensure that the comparison produced at least one revision.
        if (docOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were generated during comparison.");
        }

        // Save the comparison result (contains revisions) to a new file.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparison-result.docx");
        docOriginal.Save(resultPath, SaveFormat.Docx);
    }
}
