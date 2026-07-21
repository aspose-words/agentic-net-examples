using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string basePath = Directory.GetCurrentDirectory();
        string originalPath = Path.Combine(basePath, "Original.docx");
        string revisedPath = Path.Combine(basePath, "Revised.docx");
        string resultPath = Path.Combine(basePath, "ComparedWithRevisions.docx");

        // -----------------------------------------------------------------
        // Create the original document with some content.
        // -----------------------------------------------------------------
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a single paragraph.");
        original.Save(originalPath);

        // -----------------------------------------------------------------
        // Create the revised document with intentional differences.
        // -----------------------------------------------------------------
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document."); // changed word
        builderRevised.Writeln("It contains a single paragraph with an extra line."); // added line
        revised.Save(revisedPath);

        // -----------------------------------------------------------------
        // Load the two documents from disk.
        // -----------------------------------------------------------------
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // -----------------------------------------------------------------
        // Perform the comparison. Revisions will be added to docOriginal.
        // -----------------------------------------------------------------
        string author = "AsposeUser";
        DateTime compareTime = DateTime.Now;
        docOriginal.Compare(docRevised, author, compareTime);

        // Verify that at least one revision was created.
        if (docOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Optional: output revision count to the console.
        Console.WriteLine($"Revisions created: {docOriginal.Revisions.Count}");

        // -----------------------------------------------------------------
        // Save the document that now contains the revisions.
        // -----------------------------------------------------------------
        docOriginal.Save(resultPath);
    }
}
