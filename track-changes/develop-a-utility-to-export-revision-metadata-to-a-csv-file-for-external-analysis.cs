using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (not a revision).
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert new content – this will be recorded as an insertion revision.
        builder.Writeln("This paragraph was added while tracking changes.");

        // Delete a run from the first paragraph – this will be recorded as a deletion revision.
        // The first paragraph currently has one run ("This is the original paragraph.\r").
        // Removing it creates a deletion revision.
        if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
        {
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Export revision metadata to a CSV file.
        string csvPath = "Revisions.csv";
        using (StreamWriter writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Write CSV header.
            writer.WriteLine("RevisionType,Author,DateTime,Text");

            // Iterate through all revisions in the document.
            foreach (Revision rev in doc.Revisions)
            {
                string type = rev.RevisionType.ToString();
                string author = rev.Author;
                string date = rev.DateTime.ToString("o"); // ISO 8601 format.
                // Get the text associated with the revision; replace line breaks and commas to keep CSV integrity.
                string text = rev.ParentNode?.GetText()
                    .Replace("\r", " ")
                    .Replace("\n", " ")
                    .Replace(",", " ")
                    .Trim();

                writer.WriteLine($"{type},{author},{date},{text}");
            }
        }

        // Save the document with revisions for reference.
        doc.Save("TrackedDocument.docx");
    }
}
