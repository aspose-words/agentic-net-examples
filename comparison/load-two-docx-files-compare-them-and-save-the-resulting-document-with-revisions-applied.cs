using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Determine a folder for the temporary files.
        string workFolder = Directory.GetCurrentDirectory();

        // Paths for the two source documents and the comparison result.
        string originalPath = Path.Combine(workFolder, "Original.docx");
        string revisedPath = Path.Combine(workFolder, "Revised.docx");
        string resultPath = Path.Combine(workFolder, "ComparedWithRevisions.docx");

        // -----------------------------------------------------------------
        // Create the first document (original).
        // -----------------------------------------------------------------
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a single paragraph.");
        original.Save(originalPath);

        // -----------------------------------------------------------------
        // Create the second document (revised) with intentional differences.
        // -----------------------------------------------------------------
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document."); // Changed first line.
        builderRevised.Writeln("It now contains two paragraphs."); // Modified second line.
        builderRevised.Writeln("Additional content was added.");   // New third line.
        revised.Save(revisedPath);

        // -----------------------------------------------------------------
        // Load the documents from disk.
        // -----------------------------------------------------------------
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // -----------------------------------------------------------------
        // Compare the documents. Revisions will be added to docOriginal.
        // -----------------------------------------------------------------
        string author = "Comparer";
        DateTime compareTime = DateTime.Now;
        docOriginal.Compare(docRevised, author, compareTime);

        // Verify that at least one revision was created.
        if (docOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // -----------------------------------------------------------------
        // Save the document that now contains the revision markup.
        // -----------------------------------------------------------------
        docOriginal.Save(resultPath);
    }
}
