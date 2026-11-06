using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace RevisionUtility
{
    public class Program
    {
        // Minimum word count a revision must have to be accepted.
        private const int MinWordCount = 5;

        public static void Main()
        {
            // Create a new document and add initial content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the original paragraph of the document.");

            // Enable tracking of revisions.
            doc.StartTrackRevisions("RevisionUtility", DateTime.Now);

            // Insert a short paragraph (will be rejected).
            builder.Writeln("Short.");

            // Insert a longer paragraph (will be accepted).
            builder.Writeln("This is a longer inserted sentence with several words.");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the document that contains the revisions.
            doc.Save("Revisions.docx");

            // Create a stable list of revisions to avoid modifying the collection while iterating.
            List<Revision> revisions = doc.Revisions.Cast<Revision>().ToList();

            // Process each revision: accept only if it meets the minimum word count.
            foreach (Revision rev in revisions)
            {
                if (rev.RevisionType == RevisionType.Insertion)
                {
                    // Get the text of the inserted node.
                    string insertedText = rev.ParentNode.GetText();

                    // Count words in the inserted text.
                    int wordCount = CountWords(insertedText);

                    // Accept or reject based on the word count threshold.
                    if (wordCount >= MinWordCount)
                        rev.Accept();
                    else
                        rev.Reject();
                }
                else
                {
                    // For non‑insertion revisions, reject by default.
                    rev.Reject();
                }
            }

            // Save the final document after processing revisions.
            doc.Save("Processed.docx");
        }

        // Helper method to count words in a string.
        private static int CountWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            char[] delimiters = { ' ', '\t', '\r', '\n' };
            return text.Split(delimiters, StringSplitOptions.RemoveEmptyEntries).Length;
        }
    }
}
