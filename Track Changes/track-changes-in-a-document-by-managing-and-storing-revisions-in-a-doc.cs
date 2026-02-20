using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace RevisionDemo
{
    // Implements IRevisionCriteria to filter revisions by author and type.
    public class RevisionCriteria : IRevisionCriteria
    {
        private readonly string _author;
        private readonly RevisionType _type;

        public RevisionCriteria(string author, RevisionType type)
        {
            _author = author;
            _type = type;
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author == _author && revision.RevisionType == _type;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some initial text – this is not a revision.
            builder.Writeln("Paragraph before tracking changes.");

            // Start tracking revisions with author "Alice".
            doc.StartTrackRevisions("Alice", DateTime.Now);
            builder.Writeln("This line is an insertion revision.");
            // Stop tracking for now.
            doc.StopTrackRevisions();

            // Start a second tracking session with a different author.
            doc.StartTrackRevisions("Bob", DateTime.Now);
            // Delete the first paragraph to create a deletion revision.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            firstParagraph.Remove();
            // Insert another paragraph.
            builder.Writeln("Another inserted line by Bob.");
            doc.StopTrackRevisions();

            // At this point the document has several revisions.
            Console.WriteLine($"Total revisions: {doc.Revisions.Count}");

            // Accept only Alice's insertion revisions.
            doc.Revisions.Accept(new RevisionCriteria("Alice", RevisionType.Insertion));
            // Reject all remaining deletions (regardless of author).
            doc.Revisions.Reject(new RevisionCriteria("", RevisionType.Deletion));

            // Display remaining revisions after accept/reject.
            Console.WriteLine($"Revisions after processing: {doc.Revisions.Count}");
            foreach (Revision rev in doc.Revisions)
            {
                Console.WriteLine($"Author: {rev.Author}, Type: {rev.RevisionType}, Text: {rev.ParentNode?.GetText().Trim()}");
            }

            // Save the document with revisions preserved.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                // Ensure that revisions are kept in the saved file.
                UpdateFields = false
            };
            doc.Save("TrackedChangesDemo.docx", saveOptions);
        }
    }
}
