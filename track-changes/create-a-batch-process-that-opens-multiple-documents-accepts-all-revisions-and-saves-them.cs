using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder to store sample documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create a few sample documents with revisions.
        string[] sampleFiles = { "Doc1.docx", "Doc2.docx", "Doc3.docx" };
        foreach (string fileName in sampleFiles)
        {
            string filePath = Path.Combine(docsFolder, fileName);
            CreateDocumentWithRevisions(filePath);
        }

        // Batch process: open each document, accept all revisions, and save in place.
        foreach (string docPath in Directory.GetFiles(docsFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            doc.AcceptAllRevisions(); // Accept all tracked changes.
            doc.Save(docPath);        // Overwrite the original file.
        }
    }

    // Creates a document, adds some tracked changes, and saves it.
    private static void CreateDocumentWithRevisions(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (no revisions).
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Add content that will be recorded as a revision.
        builder.Writeln("This line is added as a revision.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document.
        doc.Save(filePath);
    }
}
