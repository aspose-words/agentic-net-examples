using System;
using System.IO;
using System.Drawing;
using Aspose.Words;

public class TrackChangesDemo
{
    public static void Main()
    {
        // Create a sample document.
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
        sampleBuilder.Writeln("This is the original paragraph.");
        sampleDoc.Save(samplePath);

        // Load the document.
        Document doc = new Document(samplePath);

        // Enable revision tracking.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Apply a formatting change (won't be recorded as a revision, but performed while tracking).
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        Run firstRun = (Run)firstParagraph.Runs[0];
        firstRun.Font.Bold = true;
        // Correct property name for color.
        firstRun.Font.Color = Color.Blue;

        // Insert new text to generate an insertion revision.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This line was added while tracking changes.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the revised document.
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        doc.Save(revisedPath);

        // List resulting revision types.
        Console.WriteLine("Revisions found in the document:");
        foreach (Revision rev in doc.Revisions)
        {
            string text = rev.ParentNode?.GetText().Trim() ?? string.Empty;
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{text}\"");
        }
    }
}
