using System;
using System.IO;
using Aspose.Words;

namespace BatchRejectRevisions
{
    // Criteria that matches revisions authored by a specific user.
    public class AuthorRevisionCriteria : IRevisionCriteria
    {
        private readonly string _authorName;

        public AuthorRevisionCriteria(string authorName)
        {
            _authorName = authorName;
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author.Equals(_authorName, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class Program
    {
        // Entry point.
        public static void Main()
        {
            // Define input and output folders.
            string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

            Directory.CreateDirectory(inputDir);
            Directory.CreateDirectory(outputDir);

            // Create sample documents with revisions from two authors.
            CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"));
            CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"));

            // Author whose revisions should be rejected.
            const string targetAuthor = "Bob";

            // Process each document in the input folder.
            foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Reject revisions authored by the target user.
                doc.Revisions.Reject(new AuthorRevisionCriteria(targetAuthor));

                // Save the processed document to the output folder.
                string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
                doc.Save(outputPath);
            }

            // Indicate completion.
            Console.WriteLine("Batch processing completed.");
        }

        // Generates a sample document containing revisions from two different authors.
        private static void CreateSampleDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Initial content without tracking.
            builder.Writeln("Original paragraph.");

            // Track revisions by Alice.
            doc.StartTrackRevisions("Alice", DateTime.Now);
            builder.Writeln("Alice's inserted paragraph.");
            doc.StopTrackRevisions();

            // Track revisions by Bob.
            doc.StartTrackRevisions("Bob", DateTime.Now);
            builder.Writeln("Bob's inserted paragraph.");
            // Delete a run to create a deletion revision by Bob.
            if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
                doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
            doc.StopTrackRevisions();

            // Save the sample document.
            doc.Save(filePath);
        }
    }
}
