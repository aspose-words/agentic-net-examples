using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace BatchAcceptRevisions
{
    public class Program
    {
        public static void Main()
        {
            // Folder to store sample documents.
            string folder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
            Directory.CreateDirectory(folder);

            // List of document file names.
            List<string> docFiles = new List<string>
            {
                Path.Combine(folder, "Document1.docx"),
                Path.Combine(folder, "Document2.docx"),
                Path.Combine(folder, "Document3.docx")
            };

            // Ensure each document exists and contains revisions.
            foreach (string filePath in docFiles)
            {
                if (!File.Exists(filePath))
                {
                    CreateSampleDocumentWithRevisions(filePath);
                }
            }

            // Batch process: open each document, accept all revisions, and save in place.
            foreach (string filePath in docFiles)
            {
                Document doc = new Document(filePath);
                // Accept all revisions in the document.
                doc.AcceptAllRevisions();
                // Save the document, overwriting the original file.
                doc.Save(filePath);
            }
        }

        // Creates a simple document, enables tracking, makes a change to generate a revision, then stops tracking.
        private static void CreateSampleDocumentWithRevisions(string filePath)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add initial content.
            builder.Writeln("Original paragraph.");

            // Start tracking revisions.
            doc.StartTrackRevisions("SampleAuthor", DateTime.Now);

            // Make a change that will be recorded as a revision.
            builder.Writeln("Inserted paragraph while tracking revisions.");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the document to the specified path.
            doc.Save(filePath);
        }
    }
}
