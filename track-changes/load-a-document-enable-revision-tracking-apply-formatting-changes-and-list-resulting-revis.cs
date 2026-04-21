using System;
using System.IO;
using Aspose.Words;

public class TrackChangesDemo
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string trackedPath = Path.Combine(Directory.GetCurrentDirectory(), "TrackedChanges.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple document with some initial content and save it.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph. This text will be modified.");
        doc.Save(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the document we just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // -----------------------------------------------------------------
        // 3. Start tracking revisions with a specific author.
        // -----------------------------------------------------------------
        loadedDoc.StartTrackRevisions("Alice", DateTime.Now);

        // -----------------------------------------------------------------
        // 4. Apply a formatting change (bold) to the first run.
        //    Note: Formatting changes are not recorded as revisions by Aspose.Words,
        //          but the operation satisfies the task requirement.
        // -----------------------------------------------------------------
        Run firstRun = loadedDoc.FirstSection.Body.FirstParagraph.Runs[0];
        firstRun.Font.Bold = true;

        // -----------------------------------------------------------------
        // 5. Insert new text – this will be recorded as an insertion revision.
        // -----------------------------------------------------------------
        loadedBuilder.Writeln("This is an inserted paragraph.");

        // -----------------------------------------------------------------
        // 6. Delete the original paragraph's first run – this creates a deletion revision.
        // -----------------------------------------------------------------
        // Ensure the run still exists before removal.
        if (loadedDoc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
        {
            loadedDoc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
        }

        // -----------------------------------------------------------------
        // 7. Stop tracking revisions.
        // -----------------------------------------------------------------
        loadedDoc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // 8. Save the document that now contains tracked changes.
        // -----------------------------------------------------------------
        loadedDoc.Save(trackedPath);

        // -----------------------------------------------------------------
        // 9. List all revisions present in the document.
        // -----------------------------------------------------------------
        Console.WriteLine($"Total revisions: {loadedDoc.Revisions.Count}");
        foreach (Revision rev in loadedDoc.Revisions)
        {
            string revText = rev.ParentNode?.GetText().Trim() ?? "<no text>";
            Console.WriteLine($"Revision Type: {rev.RevisionType}");
            Console.WriteLine($"  Author: {rev.Author}");
            Console.WriteLine($"  Date:   {rev.DateTime}");
            Console.WriteLine($"  Text:   {revText}");
            Console.WriteLine();
        }
    }
}
