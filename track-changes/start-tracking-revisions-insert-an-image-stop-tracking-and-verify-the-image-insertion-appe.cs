using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text (not a revision).
        builder.Writeln("Initial text before tracking.");

        // Start tracking revisions.
        string author = "TestAuthor";
        doc.StartTrackRevisions(author, DateTime.Now);

        // Insert a simple 1x1 PNG image while tracking is enabled.
        // This PNG is a transparent pixel encoded in base64.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=");
        builder.InsertImage(pngBytes);

        // Add more text after the image (still tracked).
        builder.Writeln("Text after image.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that an insertion revision exists for the image.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("No revisions were created.");

        // Find the first insertion revision that is a shape.
        Revision imageRevision = doc.Revisions
            .FirstOrDefault(r => r.RevisionType == RevisionType.Insertion && r.ParentNode is Shape);

        if (imageRevision == null)
            throw new InvalidOperationException("Insertion revision for the image was not found.");

        if (imageRevision.Author != author)
            throw new InvalidOperationException($"Expected author '{author}', got '{imageRevision.Author}'.");

        // The parent node of the revision should be a Shape (the inserted image) and marked as an insert revision.
        if (!(imageRevision.ParentNode is Shape shape) || !shape.IsInsertRevision)
            throw new InvalidOperationException("The inserted image is not marked as an insertion revision.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TrackChangesImage.docx");
        doc.Save(outputPath);
    }
}
