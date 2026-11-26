using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace BatchRevisionRejector
{
    public class Program
    {
        // Entry point of the console application.
        public static void Main()
        {
            // Folder that will contain sample documents.
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
            Directory.CreateDirectory(folderPath);

            // Create sample documents with revisions from different authors.
            CreateSampleDocument(Path.Combine(folderPath, "Sample1.docx"));
            CreateSampleDocument(Path.Combine(folderPath, "Sample2.docx"));

            // Author whose revisions should be rejected.
            string targetAuthor = "UserA";

            // Process each .docx file in the folder.
            foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Collect revisions to avoid modifying the collection while iterating.
                var revisions = doc.Revisions.Cast<Revision>().ToList();

                // Reject revisions authored by the target user.
                foreach (Revision rev in revisions)
                {
                    if (string.Equals(rev.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                    {
                        rev.Reject();
                    }
                }

                // Save the modified document (overwrite original).
                doc.Save(filePath);
            }

            // Optional: indicate processing is complete.
            Console.WriteLine("Revision rejection completed for author: " + targetAuthor);
        }

        // Creates a sample document with revisions from two different authors.
        private static void CreateSampleDocument(string filePath)
        {
            // Initialize a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add initial content (no revisions).
            builder.Writeln("Original paragraph.");

            // First set of revisions by UserA.
            doc.StartTrackRevisions("UserA", DateTime.Now);
            builder.Writeln("Paragraph added by UserA.");
            doc.StopTrackRevisions();

            // Second set of revisions by UserB.
            doc.StartTrackRevisions("UserB", DateTime.Now);
            builder.Writeln("Paragraph added by UserB.");
            doc.StopTrackRevisions();

            // Save the document to the specified path.
            doc.Save(filePath);
        }
    }
}
