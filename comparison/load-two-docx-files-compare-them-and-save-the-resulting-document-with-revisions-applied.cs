using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string originalPath = Path.Combine(baseDir, "Original.docx");
        string revisedPath = Path.Combine(baseDir, "Revised.docx");
        string resultPath = Path.Combine(baseDir, "ComparedResult.docx");

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world!");
        builder1.Writeln("This is the original document.");
        original.Save(originalPath);

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello world!"); // unchanged line
        builder2.Writeln("This is the revised document."); // changed line
        builder2.Writeln("Additional paragraph."); // new line
        revised.Save(revisedPath);

        // Load the two documents from disk.
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // Ensure both documents are free of revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docRevised.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. Revisions will be added to docOriginal.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (docOriginal.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the resulting document that contains the revisions.
        docOriginal.Save(resultPath);
    }
}
