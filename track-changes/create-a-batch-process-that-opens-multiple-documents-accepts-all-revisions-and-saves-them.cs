using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Directory to hold sample documents.
        string docsDir = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsDir);

        // Create sample documents with revisions.
        CreateSampleDocument(Path.Combine(docsDir, "Doc1.docx"));
        CreateSampleDocument(Path.Combine(docsDir, "Doc2.docx"));

        // Batch process: accept all revisions and save in place.
        foreach (string filePath in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(filePath);
            doc.AcceptAllRevisions();
            doc.Save(filePath);
        }
    }

    // Creates a document, enables tracking, makes a change, then saves it.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content.
        builder.Writeln("Original paragraph.");

        // Enable tracking.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make a revision: insert a new paragraph.
        builder.Writeln("Inserted paragraph while tracking changes.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Save the document.
        doc.Save(filePath);
    }
}
