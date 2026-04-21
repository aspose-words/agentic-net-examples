using System;
using System.IO;
using Aspose.Words;

public class RevisionExportUtility
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content (not a revision).
        builder.Write("This is the original paragraph. ");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Write("This text was inserted while tracking changes. ");

        // Delete the first run (the original text) – this creates a deletion revision.
        // The first run is the text added before tracking started.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document so the revisions are persisted.
        string docPath = "SampleRevisions.docx";
        doc.Save(docPath);

        // Export revision metadata to a CSV file.
        string csvPath = "RevisionsMetadata.csv";
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // Write CSV header.
            writer.WriteLine("RevisionIndex,Author,DateTime,RevisionType,Text");

            // Iterate through all revisions in the document.
            for (int i = 0; i < doc.Revisions.Count; i++)
            {
                Revision rev = doc.Revisions[i];

                // Get the text associated with the revision's parent node.
                // Some revisions (e.g., style changes) may have a null ParentNode.
                string text = rev.ParentNode != null
                    ? rev.ParentNode.GetText().Replace("\r", " ").Replace("\n", " ").Trim()
                    : string.Empty;

                // Write a CSV line with the revision details.
                writer.WriteLine($"{i},\"{rev.Author}\",{rev.DateTime:o},{rev.RevisionType},\"{text}\"");
            }
        }

        // The utility finishes execution automatically.
    }
}
