using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class RevisionLogger
{
    private readonly StringBuilder _log = new StringBuilder();

    public void Log(Revision revision)
    {
        string text = revision.ParentNode != null ? revision.ParentNode.GetText().Trim() : string.Empty;
        _log.AppendLine($"Author: {revision.Author}, Date: {revision.DateTime}, Type: {revision.RevisionType}, Text: \"{text}\"");
    }

    public void WriteToFile(string filePath)
    {
        File.WriteAllText(filePath, _log.ToString());
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content (not tracked).
        builder.Write("Original paragraph. ");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Perform edits that will be recorded as revisions.
        builder.Write("Inserted text while tracking. ");
        builder.Writeln("New line added.");

        // Delete a run to create a deletion revision.
        if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
        {
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        }

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Initialize the custom logger.
        RevisionLogger logger = new RevisionLogger();

        // Capture metadata of each revision.
        foreach (Revision rev in doc.Revisions)
        {
            logger.Log(rev);
        }

        // Save the document and the revision log.
        string docPath = "TrackedDocument.docx";
        string logPath = "RevisionLog.txt";

        doc.Save(docPath);
        logger.WriteToFile(logPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to {docPath}");
        Console.WriteLine($"Revision log saved to {logPath}");
    }
}
