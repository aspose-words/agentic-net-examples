using System;
using System.Collections.Generic;
using Aspose.Words;

public class RevisionProcessor
{
    // Minimum number of words a revision must contain to be accepted.
    private const int MinimumWordCount = 3;

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("This is the original paragraph. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Add a short revision (2 words) – should be rejected.
        builder.Writeln("Hi there. ");

        // Add a longer revision (5 words) – should be accepted.
        builder.Writeln("This is a longer revision text. ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Process each revision according to the word count threshold.
        // Collect revisions first to avoid modifying the collection while iterating.
        List<Revision> revisions = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
            revisions.Add(rev);

        foreach (Revision rev in revisions)
        {
            // Get the text associated with the revision.
            string revText = rev.ParentNode?.GetText() ?? string.Empty;
            int wordCount = CountWords(revText);

            if (wordCount >= MinimumWordCount)
                rev.Accept(); // Accept revisions meeting the threshold.
            else
                rev.Reject(); // Reject revisions that do not meet the threshold.
        }

        // Save the resulting document.
        doc.Save("ProcessedRevisions.docx");
    }

    // Helper method to count words in a given string.
    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return 0;

        // Split on whitespace characters.
        string[] words = text.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
        return words.Length;
    }
}
