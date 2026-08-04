using System;
using Aspose.Words;

namespace RevisionLoggerExample
{
    // Simple logger that records revision metadata.
    public class RevisionLogger
    {
        // Logs a single revision to the console.
        public void LogRevision(Revision revision)
        {
            // Capture basic metadata.
            string author = revision.Author;
            DateTime date = revision.DateTime;
            RevisionType type = revision.RevisionType;

            // For insertions and deletions the affected text is in ParentNode.
            string text = revision.ParentNode != null ? revision.ParentNode.GetText().Trim() : "<no text>";

            Console.WriteLine($"Revision - Author: {author}, Date: {date}, Type: {type}, Text: \"{text}\"");
        }

        // Logs all revisions in a document.
        public void LogAllRevisions(Document doc)
        {
            foreach (Revision rev in doc.Revisions)
            {
                LogRevision(rev);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write initial content (not tracked).
            builder.Writeln("Paragraph 1: Original content.");
            builder.Writeln("Paragraph 2: Original content.");

            // Start tracking revisions with a specific author.
            string author = "Alice";
            DateTime revisionStart = DateTime.Now;
            doc.StartTrackRevisions(author, revisionStart);

            // Insert new text (creates an insertion revision).
            builder.Writeln("Paragraph 3: Added while tracking.");

            // Delete a run from the first paragraph (creates a deletion revision).
            // Remove the word "Original" from the first paragraph.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            Run runToRemove = null;

            // Find the run containing the word "Original".
            foreach (Run run in firstParagraph.Runs)
            {
                if (run.Text.Contains("Original"))
                {
                    runToRemove = run;
                    break;
                }
            }

            // If found, remove it to generate a deletion revision.
            runToRemove?.Remove();

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Save the document with revisions.
            string outputPath = "TrackedDocument.docx";
            doc.Save(outputPath);

            // Log revision metadata.
            RevisionLogger logger = new RevisionLogger();
            Console.WriteLine("=== Revision Log ===");
            logger.LogAllRevisions(doc);
        }
    }
}
