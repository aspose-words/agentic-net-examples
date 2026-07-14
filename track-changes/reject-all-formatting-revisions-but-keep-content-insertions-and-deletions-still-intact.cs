using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsTrackChangesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add initial content that will later be modified.
            builder.Writeln("This is the original paragraph.");

            // Start tracking revisions.
            doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

            // ----- Insertion revision -----
            // Insert a new paragraph – this will be recorded as an insertion.
            builder.Writeln("This paragraph was inserted while tracking changes.");

            // ----- Formatting revision -----
            // Change the style of the first paragraph – this should be recorded as a format change.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            firstParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // ----- Deletion revision -----
            // Remove a run from the first paragraph – this will be recorded as a deletion.
            if (firstParagraph.Runs.Count > 0)
                firstParagraph.Runs[0].Remove();

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Reject only formatting revisions, keep insertions and deletions intact.
            foreach (Revision revision in doc.Revisions)
            {
                if (revision.RevisionType == RevisionType.FormatChange)
                    revision.Reject();
            }

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
