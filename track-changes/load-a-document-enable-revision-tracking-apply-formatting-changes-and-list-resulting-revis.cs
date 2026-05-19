using System;
using System.IO;
using Aspose.Words;

public class TrackChangesExample
{
    public static void Main()
    {
        // Define file paths for the sample documents.
        string basePath = Directory.GetCurrentDirectory();
        string originalPath = Path.Combine(basePath, "Sample.docx");
        string revisedPath = Path.Combine(basePath, "SampleWithRevisions.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a simple document and save it.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Original text. ");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document, start tracking revisions, and modify it.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);
        DocumentBuilder revBuilder = new DocumentBuilder(loadedDoc);

        // Enable revision tracking with a specific author.
        loadedDoc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert new text – this will create an insertion revision.
        revBuilder.Write("Inserted text. ");

        // Apply a formatting change (bold) to the newly inserted run.
        // Note: Formatting changes are not recorded as revisions by Aspose.Words,
        // but the operation demonstrates applying formatting while tracking is on.
        Run insertedRun = revBuilder.CurrentParagraph.Runs[revBuilder.CurrentParagraph.Runs.Count - 1];
        insertedRun.Font.Bold = true;

        // Delete the original run – this will create a deletion revision.
        Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count > 0)
        {
            firstParagraph.Runs[0].Remove();
        }

        // Stop tracking further changes.
        loadedDoc.StopTrackRevisions();

        // Save the document that now contains revisions.
        loadedDoc.Save(revisedPath);

        // -----------------------------------------------------------------
        // Step 3: List all revisions present in the document.
        // -----------------------------------------------------------------
        Console.WriteLine("Revisions found in the document:");
        foreach (Revision rev in loadedDoc.Revisions)
        {
            string revText = rev.ParentNode?.GetText().Trim() ?? "[No parent node]";
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            Console.WriteLine($"  Affected text: \"{revText}\"");
        }
    }
}
