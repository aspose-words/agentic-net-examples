using System;
using System.IO;
using Aspose.Words;

public class BatchRevisionProcessor
{
    public static void Main()
    {
        // Define sample document file names.
        string[] docFiles = { "Sample1.docx", "Sample2.docx", "Sample3.docx" };

        // Create sample documents with revisions.
        foreach (string file in docFiles)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add initial content.
            builder.Writeln("Original paragraph.");

            // Start tracking revisions.
            doc.StartTrackRevisions("BatchProcessor", DateTime.Now);

            // Make some changes that will be recorded as revisions.
            builder.Writeln("Inserted paragraph while tracking.");
            builder.MoveToDocumentEnd();
            builder.Writeln("Another inserted paragraph.");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the document to disk.
            doc.Save(file);
        }

        // Batch process: open each document, accept all revisions, and save in place.
        foreach (string file in docFiles)
        {
            // Load the existing document.
            Document doc = new Document(file);

            // Accept all revisions in the document.
            doc.AcceptAllRevisions();

            // Save the document, overwriting the original file.
            doc.Save(file);
        }
    }
}
