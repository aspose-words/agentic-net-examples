using System;
using Aspose.Words;

public class RevisionUtility
{
    public static void Main()
    {
        // Minimum number of words a revision must contain to be accepted.
        const int MinWordCount = 3;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("Original content. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Utility", DateTime.Now);

        // Insert a short revision (1 word) – should be rejected.
        builder.Writeln("Hi.");

        // Insert a longer revision (6 words) – should be accepted.
        builder.Writeln("This is a longer inserted sentence.");

        // Delete a run to create a deletion revision (will be rejected by default).
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Process revisions: accept only those that meet the word count threshold.
        // Iterate backwards because accepting/rejecting modifies the collection.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];

            // Only evaluate insertion revisions; other types are rejected.
            if (rev.RevisionType == RevisionType.Insertion)
            {
                // Get the text of the revised node and count words.
                string text = rev.ParentNode.GetText().Trim();
                int wordCount = text.Split(new[] { ' ', '\t', '\r', '\n' },
                                            StringSplitOptions.RemoveEmptyEntries).Length;

                if (wordCount >= MinWordCount)
                    rev.Accept();
                else
                    rev.Reject();
            }
            else
            {
                // Reject all non‑insertion revisions.
                rev.Reject();
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
