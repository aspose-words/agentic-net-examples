using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Start tracking revisions with a specific author and timestamp.
        string author = "Alice";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Use DocumentBuilder to make changes that will generate revisions.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an original paragraph.
        builder.Writeln("Original paragraph.");

        // Insert a new paragraph (this will be an insertion revision).
        builder.Writeln("Inserted paragraph.");

        // Change formatting of the inserted paragraph (format change revision).
        Paragraph insertedParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        builder.MoveTo(insertedParagraph);
        builder.Font.Bold = true;

        // Delete the original paragraph (deletion revision).
        Paragraph originalParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        originalParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document with revisions.
        string docPath = "TrackedDocument.docx";
        doc.Save(docPath);

        // Prepare a log file for audit.
        string logPath = "RevisionLog.txt";
        using (StreamWriter logWriter = new StreamWriter(logPath, false))
        {
            // Iterate through all revisions and capture metadata.
            foreach (Revision rev in doc.Revisions)
            {
                string revType = rev.RevisionType.ToString();
                string revAuthor = rev.Author;
                DateTime revDate = rev.DateTime;
                string revText = rev.ParentNode?.GetText().Trim() ?? "<no text>";

                string logEntry = $"Type: {revType}, Author: {revAuthor}, Date: {revDate:u}, Text: \"{revText}\"";
                Console.WriteLine(logEntry);
                logWriter.WriteLine(logEntry);
            }
        }

        // Indicate completion.
        Console.WriteLine("Revision audit completed. Document saved to " + Path.GetFullPath(docPath));
        Console.WriteLine("Log saved to " + Path.GetFullPath(logPath));
    }
}
