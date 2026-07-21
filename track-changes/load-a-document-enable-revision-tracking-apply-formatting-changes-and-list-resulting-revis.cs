using System;
using System.IO;
using Aspose.Words;

public class TrackChangesExample
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        string trackedPath = Path.Combine(Directory.GetCurrentDirectory(), "Tracked.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample document and save it.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the original paragraph.");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document, start tracking revisions, apply changes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);
        loadedDoc.StartTrackRevisions("Alice", DateTime.Now);

        // Apply a formatting change (bold) to the first run.
        // Formatting changes are not recorded as revisions, but the operation is performed.
        Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;
        Run firstRun = (Run)firstParagraph.Runs[0];
        firstRun.Font.Bold = true;

        // Insert new text while tracking to generate an insertion revision.
        DocumentBuilder trackingBuilder = new DocumentBuilder(loadedDoc);
        trackingBuilder.Writeln("Inserted paragraph while tracking.");

        // Stop tracking further changes.
        loadedDoc.StopTrackRevisions();

        // Save the document with revisions.
        loadedDoc.Save(trackedPath);

        // -----------------------------------------------------------------
        // Step 3: List all revisions and their types.
        // -----------------------------------------------------------------
        Console.WriteLine("Revisions found in the document:");
        foreach (Revision rev in loadedDoc.Revisions)
        {
            string revText = rev.ParentNode?.GetText().Trim() ?? string.Empty;
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{revText}\"");
        }
    }
}
