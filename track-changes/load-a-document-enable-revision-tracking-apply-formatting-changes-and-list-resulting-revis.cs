using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace TrackChangesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the sample documents.
            string originalPath = "Sample.docx";
            string revisedPath = "Revised.docx";

            // 1. Create a simple document with initial content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample paragraph for revision tracking.");
            doc.Save(originalPath);

            // 2. Load the document that was just created.
            Document loadedDoc = new Document(originalPath);

            // 3. Enable revision tracking.
            string author = "John Doe";
            DateTime revisionDate = DateTime.Now;
            loadedDoc.StartTrackRevisions(author, revisionDate);

            // 4. Apply changes while tracking is enabled.

            // a) Change the style of the first paragraph (format revision).
            Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;
            firstParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // b) Make all runs in the first paragraph bold (format revision).
            foreach (Run run in firstParagraph.GetChildNodes(NodeType.Run, true))
            {
                run.Font.Bold = true;
            }

            // c) Insert a new paragraph (insertion revision) to guarantee at least one revision.
            DocumentBuilder editBuilder = new DocumentBuilder(loadedDoc);
            editBuilder.MoveTo(firstParagraph);
            editBuilder.Writeln("Additional text added while tracking is on.");

            // 5. Stop tracking revisions.
            loadedDoc.StopTrackRevisions();

            // 6. Save the revised document.
            loadedDoc.Save(revisedPath);

            // 7. List the resulting revisions (type, author, and date).
            RevisionCollection revisions = loadedDoc.Revisions;
            if (revisions.Count == 0)
                throw new InvalidOperationException("No revisions were generated.");

            Console.WriteLine("Revisions found:");
            foreach (Revision rev in revisions)
            {
                // Output revision type, author and date.
                Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            }
        }
    }
}
