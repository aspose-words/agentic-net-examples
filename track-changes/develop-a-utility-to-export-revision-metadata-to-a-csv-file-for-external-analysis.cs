using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will not be a revision.
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Delete the first run of the original paragraph – this creates a deletion revision.
        if (doc.FirstSection.Body.Paragraphs.Count > 0 && doc.FirstSection.Body.Paragraphs[0].Runs.Count > 0)
        {
            doc.FirstSection.Body.Paragraphs[0].Runs[0].Remove();
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates that revisions are persisted).
        string docPath = "RevisionsDemo.docx";
        doc.Save(docPath);

        // Export revision metadata to a CSV file.
        string csvPath = "RevisionsMetadata.csv";
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // Write CSV header.
            writer.WriteLine("Index,Author,DateTime,RevisionType,Text");

            int index = 0;
            foreach (Revision rev in doc.Revisions)
            {
                // Retrieve the text associated with the revision, handling possible nulls.
                string text = rev.ParentNode != null
                    ? rev.ParentNode.GetText().Replace("\r", " ").Replace("\n", " ").Trim()
                    : string.Empty;

                // Write a CSV line with proper escaping.
                writer.WriteLine($"{index},{EscapeCsv(rev.Author)},{rev.DateTime:o},{rev.RevisionType},{EscapeCsv(text)}");
                index++;
            }
        }
    }

    // Helper method to escape CSV fields that contain commas, quotes, or line breaks.
    private static string EscapeCsv(string value)
    {
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n"))
        {
            string escaped = value.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return value;
    }
}
