using System;
using System.IO;
using Aspose.Words;

namespace BatchRejectRevisions
{
    // Criteria that matches revisions authored by a specific user.
    class RevisionAuthorCriteria : IRevisionCriteria
    {
        private readonly string _author;

        public RevisionAuthorCriteria(string author)
        {
            _author = author;
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author == _author;
        }
    }

    public class Program
    {
        // Entry point.
        public static void Main()
        {
            // Folder that will contain sample documents.
            string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
            Directory.CreateDirectory(docsFolder);

            // Create sample documents with revisions from two authors.
            CreateSampleDocument(Path.Combine(docsFolder, "Document1.docx"));
            CreateSampleDocument(Path.Combine(docsFolder, "Document2.docx"));
            CreateSampleDocument(Path.Combine(docsFolder, "Document3.docx"));

            // Author whose revisions should be rejected.
            string targetAuthor = "Alice";

            // Process each .docx file in the folder.
            foreach (string filePath in Directory.GetFiles(docsFolder, "*.docx"))
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Reject all revisions authored by the target user.
                doc.Revisions.Reject(new RevisionAuthorCriteria(targetAuthor));

                // Save the modified document (overwrite original).
                doc.Save(filePath);
            }
        }

        // Generates a document with revisions from two different authors.
        private static void CreateSampleDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Initial content (no revisions).
            builder.Writeln("Base content of the document.");

            // Revisions by Alice.
            doc.StartTrackRevisions("Alice", DateTime.Now);
            builder.Writeln("Alice's inserted paragraph.");
            doc.StopTrackRevisions();

            // Revisions by Bob.
            doc.StartTrackRevisions("Bob", DateTime.Now);
            builder.Writeln("Bob's inserted paragraph.");

            // Delete a run to create a deletion revision.
            Node firstParagraph = doc.FirstSection.Body.FirstParagraph;
            if (firstParagraph is Paragraph para && para.Runs.Count > 0)
                para.Runs[0].Remove();

            doc.StopTrackRevisions();

            // Save the sample document.
            doc.Save(filePath);
        }
    }
}
