using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document locally.
        // -----------------------------------------------------------------
        var sampleDoc = new Document();
        var builder = new DocumentBuilder(sampleDoc);
        // Write an initial line of text.
        builder.Writeln("Hello World!");
        const string samplePath = "Sample.docx";
        sampleDoc.Save(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the document and start tracking revisions.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(samplePath);
        loadedDoc.StartTrackRevisions("AuthorName", DateTime.Now);

        // -----------------------------------------------------------------
        // 3. Perform a change while tracking is enabled.
        //    Adding a new paragraph generates an Insertion revision.
        // -----------------------------------------------------------------
        var revBuilder = new DocumentBuilder(loadedDoc);
        revBuilder.Writeln("This line was added while tracking revisions.");

        // -----------------------------------------------------------------
        // 4. Stop tracking and save the revised document.
        // -----------------------------------------------------------------
        loadedDoc.StopTrackRevisions();
        const string revisedPath = "Revised.docx";
        loadedDoc.Save(revisedPath);

        // -----------------------------------------------------------------
        // 5. List all revisions with their type, author and date.
        // -----------------------------------------------------------------
        foreach (Revision rev in loadedDoc.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }

        // -----------------------------------------------------------------
        // 6. Validate that at least one revision was generated.
        // -----------------------------------------------------------------
        if (loadedDoc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");
    }
}
