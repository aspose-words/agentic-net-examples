using System;
using Aspose.Words;

public class RevisionLogger
{
    // Logs revision metadata to the console.
    public void LogRevisions(Document doc)
    {
        foreach (Revision rev in doc.Revisions)
        {
            string type = rev.RevisionType.ToString();
            string author = rev.Author;
            DateTime date = rev.DateTime;
            string text = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;

            Console.WriteLine($"Revision: Type={type}, Author={author}, Date={date}, Text=\"{text}\"");
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (will later be deleted to generate a deletion revision).
        builder.Writeln("Original paragraph.");

        // Enable tracking of revisions.
        string author = "John Doe";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Insert a new paragraph (insertion revision).
        builder.Writeln("Inserted paragraph.");

        // Delete the original paragraph (deletion revision).
        Node firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.Remove();

        // Change formatting of a new paragraph (format change revision).
        builder.Font.Bold = true;
        builder.Writeln("Bold paragraph.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document with revisions.
        doc.Save("TrackedDocument.docx");

        // Capture and output revision metadata.
        RevisionLogger logger = new RevisionLogger();
        logger.LogRevisions(doc);
    }
}
