using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class BatchAcceptRevisions
{
    public static void Main()
    {
        // Directory to store sample documents.
        string docsDir = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(docsDir);

        // List of document file names to process.
        List<string> docFiles = new List<string>
        {
            Path.Combine(docsDir, "Document1.docx"),
            Path.Combine(docsDir, "Document2.docx"),
            Path.Combine(docsDir, "Document3.docx")
        };

        // Create sample documents with revisions.
        foreach (string filePath in docFiles)
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write initial content (no revision).
            builder.Writeln("Original content.");

            // Start tracking revisions to generate changes.
            doc.StartTrackRevisions("BatchUser", DateTime.Now);

            // Add revised content.
            builder.Writeln("Added line 1.");
            builder.Writeln("Added line 2.");

            // Stop tracking.
            doc.StopTrackRevisions();

            // Save the document (creates the file).
            doc.Save(filePath);
        }

        // Batch process: open each document, accept all revisions, and save in place.
        foreach (string filePath in docFiles)
        {
            // Load the existing document.
            Document doc = new Document(filePath);

            // Accept all tracked changes.
            doc.AcceptAllRevisions();

            // Save back to the same file (overwrites original).
            doc.Save(filePath);
        }
    }
}
