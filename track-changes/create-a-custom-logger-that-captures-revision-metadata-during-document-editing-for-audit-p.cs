using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class RevisionLogger
{
    private readonly List<string> _entries = new List<string>();

    public void Log(Revision revision)
    {
        string text = revision.ParentNode?.GetText()?.Trim() ?? string.Empty;
        string entry = $"Author: {revision.Author}, " +
                       $"Date: {revision.DateTime:u}, " +
                       $"Type: {revision.RevisionType}, " +
                       $"Text: \"{text}\"";
        _entries.Add(entry);
    }

    public void Save(string filePath)
    {
        File.WriteAllLines(filePath, _entries);
    }

    public void Print()
    {
        foreach (var entry in _entries)
            Console.WriteLine(entry);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (not tracked).
        builder.Writeln("Original paragraph. ");

        // Start tracking revisions with a specific author and timestamp.
        string author = "Alice";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Make some changes that will be recorded as revisions.
        builder.Writeln("Inserted paragraph by Alice. ");
        // Delete a run to create a deletion revision.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count > 0)
            firstParagraph.Runs[0].Remove();

        // Change formatting (will not be tracked as a revision by Aspose.Words, but included for completeness).
        builder.Font.Bold = true;
        builder.Writeln("Bold paragraph added. ");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Create a logger and capture all revision metadata.
        RevisionLogger logger = new RevisionLogger();
        foreach (Revision rev in doc.Revisions)
        {
            logger.Log(rev);
        }

        // Output log to console.
        logger.Print();

        // Save the document and the revision log.
        string outputDoc = "TrackedDocument.docx";
        string logFile = "RevisionLog.txt";
        doc.Save(outputDoc);
        logger.Save(logFile);
    }
}
