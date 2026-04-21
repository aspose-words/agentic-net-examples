using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample and result documents.
        string samplePath = "Sample.docx";
        string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a blank document and save it to disk (bootstrap).
        // -----------------------------------------------------------------
        Document blankDoc = new Document();
        blankDoc.Save(samplePath);

        // -------------------------------------------------
        // Step 2: Load the previously saved blank document.
        // -------------------------------------------------
        Document doc = new Document(samplePath);

        // -------------------------------------------------
        // Step 3: Start tracking revisions.
        // -------------------------------------------------
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // -------------------------------------------------
        // Step 4: Insert a simple 2‑cell table while tracking.
        // -------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // -------------------------------------------------
        // Step 5: Stop tracking revisions.
        // -------------------------------------------------
        doc.StopTrackRevisions();

        // -------------------------------------------------
        // Step 6: Verify that a revision was created.
        // -------------------------------------------------
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // -------------------------------------------------
        // Step 7: Accept the table insertion revision.
        // -------------------------------------------------
        // In this scenario there is only one revision (the table insertion),
        // so AcceptAllRevisions safely accepts it.
        doc.AcceptAllRevisions();

        // -------------------------------------------------
        // Step 8: Save the final document.
        // -------------------------------------------------
        doc.Save(resultPath);

        // Optional: indicate completion (no user interaction required).
        Console.WriteLine($"Document processed. Output saved to '{resultPath}'.");
    }
}
