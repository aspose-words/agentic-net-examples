using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class RevisionExportUtility
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will not be a revision.
        builder.Writeln("Original content.");

        // Enable tracking of revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Perform edits that will be recorded as revisions.
        builder.Writeln("Added line 1.");
        builder.Writeln("Added line 2.");

        // Create a deletion revision by removing the first run.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates lifecycle usage).
        doc.Save("RevisionsDemo.docx");

        // Export revision metadata to a CSV file.
        ExportRevisionsToCsv(doc, "RevisionsMetadata.csv");
    }

    private static void ExportRevisionsToCsv(Document doc, string csvPath)
    {
        var sb = new StringBuilder();

        // CSV header.
        sb.AppendLine("Index,RevisionType,Author,DateTime,Text");

        int index = 0;
        foreach (Revision rev in doc.Revisions)
        {
            // Retrieve the text associated with the revision, handling possible nulls.
            string text = rev.ParentNode != null
                ? rev.ParentNode.GetText().Replace("\r", " ").Replace("\n", " ").Trim()
                : string.Empty;

            // Escape commas in the text field.
            if (text.Contains(","))
                text = $"\"{text}\"";

            // Append a CSV line with ISO 8601 date format for consistency.
            sb.AppendLine($"{index},{rev.RevisionType},{rev.Author},{rev.DateTime:o},{text}");
            index++;
        }

        // Write the CSV content to the specified file.
        File.WriteAllText(csvPath, sb.ToString());
    }
}
