using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class RevisionUtility
{
    public static void Main()
    {
        // Minimum number of words a revision must contain to be accepted.
        const int MinWordCount = 5;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("Original paragraph without tracking.");

        // Start tracking revisions.
        doc.StartTrackRevisions("UtilityAuthor", DateTime.Now);

        // Insert a short revision (less than the threshold).
        builder.Writeln("Short rev.");

        // Insert a longer revision (meets the threshold).
        builder.Writeln("This is a longer revision that should be accepted.");

        // Delete a short piece of text to create a deletion revision.
        Paragraph paraToDelete = doc.FirstSection.Body.Paragraphs[1]; // "Short rev."
        paraToDelete.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Process revisions: accept only those with word count >= MinWordCount.
        List<Revision> revisions = doc.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisions)
        {
            // Get the text associated with the revision.
            string revText = rev.ParentNode?.GetText() ?? string.Empty;

            // Count words by splitting on whitespace.
            int wordCount = revText.Split(new char[] { ' ', '\t', '\r', '\n' },
                                          StringSplitOptions.RemoveEmptyEntries).Length;

            if (wordCount >= MinWordCount)
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("RevisionsProcessed.docx");
    }
}
