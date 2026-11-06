using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample and cleaned documents.
        string originalPath = "Original.docx";
        string cleanedPath = "Cleaned.docx";

        // -----------------------------------------------------------------
        // Create a sample document with revisions.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Initial content (no revisions).
        builder.Writeln("First paragraph.");

        // Enable tracking to generate revisions.
        sampleDoc.StartTrackRevisions("SampleAuthor", DateTime.Now);

        // Content added while tracking is enabled – this will be a revision.
        builder.Writeln("Second paragraph added as a revision.");

        // Disable tracking.
        sampleDoc.StopTrackRevisions();

        // Save the document that contains revisions.
        sampleDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // Load the document, accept all revisions, and save the cleaned version.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);
        loadedDoc.AcceptAllRevisions();
        loadedDoc.Save(cleanedPath);
    }
}
