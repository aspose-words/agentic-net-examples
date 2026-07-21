using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Minimum number of words a revision must contain to be accepted.
        const int MinWordCount = 3;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not tracked as a revision.
        builder.Writeln("Original content. ");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("UtilityAuthor", DateTime.Now);

        // Insert revisions with varying word counts.
        builder.Writeln("Short.");                     // 1 word
        builder.Writeln("This is a longer revision."); // 5 words
        builder.Writeln("Another short");               // 2 words
        builder.Writeln("Accept this revision because it has enough words."); // 9 words

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Process revisions: accept only those meeting the minimum word count.
        // Iterate backwards because accepting/rejecting modifies the collection.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];

            // Consider only insertion revisions (they contain the added text).
            if (rev.RevisionType == RevisionType.Insertion)
            {
                string text = rev.ParentNode.GetText();

                // Count words by splitting on whitespace.
                int wordCount = 0;
                foreach (string word in text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    wordCount++;
                }

                if (wordCount >= MinWordCount)
                    rev.Accept(); // Accept revision meeting the threshold.
                else
                    rev.Reject(); // Reject revision that does not meet the threshold.
            }
        }

        // Save the resulting document.
        doc.Save("RevisionsFiltered.docx");
    }
}
