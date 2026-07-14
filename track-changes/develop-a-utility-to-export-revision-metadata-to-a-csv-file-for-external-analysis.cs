using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class RevisionExporter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial content that will not be tracked.
        builder.Writeln("Original paragraph. ");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert a new paragraph – this will be recorded as an insertion revision.
        builder.Writeln("Inserted paragraph while tracking. ");

        // Delete the first run (the word "Original") to create a deletion revision.
        Run firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates that revisions are persisted).
        doc.Save("TrackedDocument.docx");

        // Export revision metadata to a CSV file.
        string csvPath = "Revisions.csv";
        using (StreamWriter writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Write CSV header.
            writer.WriteLine("Author,DateTime,RevisionType,Text");

            // Iterate through each revision and write its details.
            foreach (Revision rev in doc.Revisions)
            {
                // Get the text associated with the revision's parent node, if any.
                string text = rev.ParentNode != null
                    ? rev.ParentNode.GetText()
                        .Replace("\r", " ")
                        .Replace("\n", " ")
                        .Replace(",", " ")
                    : string.Empty;

                // Write a CSV line with author, ISO 8601 date, revision type, and text.
                writer.WriteLine($"{rev.Author},{rev.DateTime:o},{rev.RevisionType},{text}");
            }
        }
    }
}
