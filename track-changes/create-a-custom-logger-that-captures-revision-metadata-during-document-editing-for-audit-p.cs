using System;
using System.Collections.Generic;
using Aspose.Words;

public class RevisionLogger
{
    private readonly List<string> _entries = new List<string>();

    public void Log(Revision revision)
    {
        string entry = $"Author: {revision.Author}, " +
                       $"Date: {revision.DateTime:u}, " +
                       $"Type: {revision.RevisionType}, " +
                       $"Text: \"{revision.ParentNode?.GetText().Trim()}\"";
        _entries.Add(entry);
    }

    public void Print()
    {
        foreach (var entry in _entries)
        {
            Console.WriteLine(entry);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (not tracked).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions.
        string author = "Alice";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Insert a new paragraph (creates an insertion revision).
        builder.Writeln("Inserted paragraph.");

        // Delete the original paragraph (creates a deletion revision).
        Paragraph originalParagraph = doc.FirstSection.Body.Paragraphs[0];
        originalParagraph.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document with revisions.
        string outputPath = "TrackedDocument.docx";
        doc.Save(outputPath);

        // Log revision metadata.
        RevisionLogger logger = new RevisionLogger();
        foreach (Revision rev in doc.Revisions)
        {
            logger.Log(rev);
        }

        // Output the log to the console.
        Console.WriteLine("Revision Audit Log:");
        logger.Print();
    }
}
