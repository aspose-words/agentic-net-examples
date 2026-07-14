using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that will not be tracked as a revision.
        builder.Writeln("Original text.");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Add new paragraphs – these will be recorded as insertion revisions.
        builder.Writeln("Added line 1.");
        builder.Writeln("Added line 2.");

        // Delete a run to generate a deletion revision.
        // Here we remove the first run of the first paragraph.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionsDemo.docx");
        doc.Save(outputPath);

        // Iterate through all revisions and log author and timestamp.
        foreach (Revision revision in doc.Revisions)
        {
            Console.WriteLine($"Author: {revision.Author}, DateTime: {revision.DateTime}");
        }
    }
}
