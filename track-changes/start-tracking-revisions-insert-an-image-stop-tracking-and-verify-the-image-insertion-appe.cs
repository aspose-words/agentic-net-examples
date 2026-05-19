using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("Document before image insertion.");

        // Start tracking revisions with a specific author.
        string author = "Test Author";
        doc.StartTrackRevisions(author, DateTime.Now);

        // Insert a simple 1x1 pixel PNG image using a base64 string.
        // This avoids the need for System.Drawing types.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(pngBase64);

        // Insert the image while tracking is enabled.
        Shape insertedShape = builder.InsertImage(imageBytes);

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (optional for visual verification).
        string outputPath = "TrackChangesImage.docx";
        doc.Save(outputPath);

        // Verify that a revision was created for the image insertion.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("No revisions were recorded.");

        Revision revision = doc.Revisions[0];
        if (revision.RevisionType != RevisionType.Insertion)
            throw new InvalidOperationException($"Expected an insertion revision, but got {revision.RevisionType}.");

        if (revision.Author != author)
            throw new InvalidOperationException($"Expected revision author '{author}', but got '{revision.Author}'.");

        // Ensure the inserted shape is marked as an insert revision.
        if (insertedShape == null)
            throw new InvalidOperationException("Inserted shape not found.");

        if (!insertedShape.IsInsertRevision)
            throw new InvalidOperationException("The inserted image is not marked as an insert revision.");

        // If all checks pass, indicate success.
        Console.WriteLine("Image insertion recorded as a revision successfully.");
    }
}
