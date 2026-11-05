using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace BatchRejectRevisions
{
    public class Program
    {
        // Entry point
        public static void Main()
        {
            // Folder that will contain sample documents
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
            Directory.CreateDirectory(folderPath);

            // Create sample documents with revisions from two different authors
            CreateSampleDocument(Path.Combine(folderPath, "Sample1.docx"));
            CreateSampleDocument(Path.Combine(folderPath, "Sample2.docx"));

            // Author whose revisions should be rejected
            string targetAuthor = "UserA";

            // Process each .docx file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
            {
                // Load the document
                Document doc = new Document(filePath);

                // Collect revisions authored by the target user
                var revisionsToReject = doc.Revisions
                                            .Where(r => string.Equals(r.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                                            .ToList();

                // Reject each matching revision
                foreach (Revision rev in revisionsToReject)
                {
                    rev.Reject();
                }

                // Save the modified document (overwrite original)
                doc.Save(filePath);
            }
        }

        // Creates a sample document with revisions from two authors
        private static void CreateSampleDocument(string filePath)
        {
            // Start with a clean document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Initial content (no revisions)
            builder.Writeln("Original paragraph.");

            // First set of revisions by UserA
            doc.StartTrackRevisions("UserA", DateTime.Now);
            builder.Writeln("Paragraph added by UserA.");
            doc.StopTrackRevisions();

            // Second set of revisions by UserB
            doc.StartTrackRevisions("UserB", DateTime.Now);
            builder.Writeln("Paragraph added by UserB.");
            doc.StopTrackRevisions();

            // Save the document
            doc.Save(filePath);
        }
    }
}
