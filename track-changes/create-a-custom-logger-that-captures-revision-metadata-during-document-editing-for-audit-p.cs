using System;
using System.Collections.Generic;
using Aspose.Words;

public class RevisionLogger
{
    public List<string> Capture(Document doc)
    {
        var logs = new List<string>();
        foreach (Revision rev in doc.Revisions)
        {
            string type = rev.RevisionType.ToString();
            string author = rev.Author;
            DateTime date = rev.DateTime;
            string text = GetRevisionText(rev);
            logs.Add($"Revision: Type={type}, Author={author}, Date={date:u}, Text=\"{text}\"");
        }
        return logs;
    }

    private string GetRevisionText(Revision rev)
    {
        // The ParentNode holds the node that was inserted, deleted or modified.
        // Trim to remove extra whitespace and line breaks.
        return rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (no revisions).
        builder.Writeln("Original paragraph.");
        builder.Writeln("Second paragraph.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // 1. Insert a new paragraph.
        builder.Writeln("Inserted paragraph while tracking.");

        // 2. Delete the second paragraph.
        Node secondParagraph = doc.GetChild(NodeType.Paragraph, 1, true);
        secondParagraph?.Remove();

        // 3. Change formatting of the first paragraph (make it bold).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        if (firstParagraph != null && firstParagraph.Runs.Count > 0)
        {
            firstParagraph.Runs[0].Font.Bold = true;
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Capture and output revision metadata.
        RevisionLogger logger = new RevisionLogger();
        List<string> revisionLogs = logger.Capture(doc);
        foreach (string log in revisionLogs)
        {
            Console.WriteLine(log);
        }

        // Save the document with revisions.
        doc.Save("TrackedDocument.docx");
    }
}
