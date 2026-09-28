using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with some initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        builder.Writeln("This is a sample paragraph.");

        // Enable revision tracking.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Apply a formatting change: make the first run bold (creates a FormatChange revision).
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        Run firstRun = (Run)firstParagraph.GetChildNodes(NodeType.Run, true)[0];
        firstRun.Font.Bold = true;

        // Insert a new paragraph (creates an Insertion revision).
        builder.MoveToDocumentEnd();
        builder.Writeln("Inserted paragraph while tracking.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // List all revisions with their types, authors, and dates.
        Console.WriteLine($"Total revisions: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }

        // Save the document (optional, demonstrates that revisions are persisted).
        doc.Save("TrackedDocument.docx");
    }
}
