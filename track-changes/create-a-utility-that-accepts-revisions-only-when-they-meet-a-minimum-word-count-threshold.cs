using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Minimum number of words a revision must contain to be accepted.
        const int minWordCount = 5;

        // Create a new blank document and a builder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not tracked as a revision.
        builder.Writeln("Original content. ");

        // Start tracking revisions with a given author.
        doc.StartTrackRevisions("UtilityAuthor", DateTime.Now);

        // Revision 1 – short text (should be rejected).
        builder.Writeln("Short.");

        // Revision 2 – longer text (should be accepted).
        builder.Writeln("This revision contains enough words to be accepted.");

        // Revision 3 – another short text (should be rejected).
        builder.Writeln("Tiny.");

        // Revision 4 – long text (should be accepted).
        builder.Writeln("Another acceptable revision with sufficient word count.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Create a snapshot of the revisions because accepting/rejecting
        // modifies the original collection.
        Revision[] revisions = doc.Revisions.Cast<Revision>().ToArray();

        foreach (Revision rev in revisions)
        {
            // Process only insertion revisions; other types are left unchanged.
            if (rev.RevisionType == RevisionType.Insertion)
            {
                // Get the text of the revision's parent node.
                string revisionText = rev.ParentNode.GetText();

                // Count words by splitting on whitespace and discarding empty entries.
                int wordCount = revisionText
                    .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Length;

                // Accept or reject based on the word count threshold.
                if (wordCount >= minWordCount)
                    rev.Accept();
                else
                    rev.Reject();
            }
        }

        // Save the resulting document.
        doc.Save("FilteredRevisions.docx");
    }
}
