using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text (not a revision).
        builder.Writeln("Initial content before tracking.");

        // Start tracking revisions with a specific author.
        string author = "Test Author";
        doc.StartTrackRevisions(author, DateTime.Now);

        // Insert a simple in‑memory PNG image (1x1 red pixel) while tracking is enabled.
        // The image is represented as a base64‑encoded byte array to avoid System.Drawing dependencies.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXK5wAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        builder.InsertImage(imageBytes);

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document (required by the rules).
        string outputPath = "TrackedRevisions.docx";
        doc.Save(outputPath);

        // Verify that the image insertion was recorded as a revision.
        bool hasInsertionRevision = doc.Revisions
            .Any(r => r.RevisionType == RevisionType.Insertion && r.ParentNode is Shape);

        if (!hasInsertionRevision)
            throw new Exception("Image insertion revision was not found.");

        // Additional check: the inserted shape should report IsInsertRevision = true.
        Shape insertedShape = doc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .FirstOrDefault();

        if (insertedShape == null || !insertedShape.IsInsertRevision)
            throw new Exception("Inserted shape does not have InsertRevision flag.");

        // If execution reaches this point, the revision was correctly created.
        Console.WriteLine("Image insertion recorded as a revision successfully.");
    }
}
