using System;
using System.Collections.Generic;
using Aspose.Words;

namespace RevisionLoggerExample
{
    // Simple logger that records revision metadata.
    public class RevisionLogger
    {
        private readonly List<string> _entries = new List<string>();

        public void Log(Revision revision)
        {
            string entry = $"Author: {revision.Author}, " +
                           $"Date: {revision.DateTime:u}, " +
                           $"Type: {revision.RevisionType}, " +
                           $"Text: \"{revision.ParentNode?.GetText().Trim()}\"";
            _entries.Add(entry);
        }

        public void WriteLog()
        {
            Console.WriteLine("=== Revision Audit Log ===");
            foreach (var entry in _entries)
            {
                Console.WriteLine(entry);
            }
            Console.WriteLine("=== End of Log ===");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Initial content (not tracked).
            builder.Writeln("Original paragraph. ");

            // Start tracking revisions with a specific author.
            doc.StartTrackRevisions("Alice", DateTime.Now);

            // Insert a new paragraph (creates an insertion revision).
            builder.Writeln("Inserted paragraph while tracking. ");

            // Delete the original paragraph (creates a deletion revision).
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            firstParagraph.Remove();

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Save the document (optional for inspection).
            doc.Save("TrackedDocument.docx");

            // Capture revision metadata.
            RevisionLogger logger = new RevisionLogger();
            foreach (Revision rev in doc.Revisions)
            {
                logger.Log(rev);
            }

            // Output the audit log.
            logger.WriteLog();
        }
    }
}
