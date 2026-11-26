using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Minimum word count required for a revision to be accepted
        const int MinimumWordCount = 5;

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (not tracked)
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Insert a short revision (1 word)
        builder.Writeln("Short");

        // Insert a longer revision (more than the threshold)
        builder.Writeln("This is a longer inserted paragraph with several words.");

        // Stop tracking revisions
        doc.StopTrackRevisions();

        // Make a copy of the revisions collection to avoid modification during enumeration
        var revisions = doc.Revisions.Cast<Revision>().ToList();

        // Process each revision
        foreach (Revision rev in revisions)
        {
            // Only consider insertions for this example
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode != null)
            {
                // Get the text of the inserted node
                string insertedText = rev.ParentNode.GetText();

                // Count words in the inserted text
                int wordCount = CountWords(insertedText);

                // Accept or reject based on the word count threshold
                if (wordCount >= MinimumWordCount)
                {
                    rev.Accept();
                }
                else
                {
                    rev.Reject();
                }
            }
        }

        // Save the resulting document
        doc.Save("Result.docx");
    }

    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return 0;

        char[] delimiters = new char[] { ' ', '\t', '\r', '\n' };
        return text.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
    }
}
