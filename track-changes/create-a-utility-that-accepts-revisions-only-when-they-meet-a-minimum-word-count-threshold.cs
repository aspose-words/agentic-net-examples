using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

public class RevisionUtility
{
    // Minimum word count a revision must have to be accepted.
    private const int MinWordCount = 3;

    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Insert a short revision (should be rejected).
        builder.Writeln("Hi.");

        // Insert a longer revision (should be accepted).
        builder.Writeln("This is a longer inserted sentence with several words.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Process revisions: accept only those meeting the word count threshold.
        int accepted = 0;
        int rejected = 0;

        // Create a snapshot of the revisions collection to avoid modification issues during iteration.
        List<Revision> revisions = doc.Revisions.Cast<Revision>().ToList();

        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion)
            {
                // Get the text of the inserted node.
                string insertedText = rev.ParentNode.GetText();

                // Count words (simple split on whitespace).
                int wordCount = CountWords(insertedText);

                if (wordCount >= MinWordCount)
                {
                    rev.Accept();
                    accepted++;
                }
                else
                {
                    rev.Reject();
                    rejected++;
                }
            }
            else
            {
                // For non‑insertion revisions, reject by default.
                rev.Reject();
                rejected++;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Output summary (no interactive input required).
        Console.WriteLine($"Revisions processed. Accepted: {accepted}, Rejected: {rejected}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Helper method to count words in a string.
    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return 0;

        // Split on whitespace characters.
        string[] words = text.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
        return words.Length;
    }
}
