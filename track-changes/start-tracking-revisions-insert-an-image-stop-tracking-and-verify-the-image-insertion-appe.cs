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

        // Add some initial text (not tracked).
        builder.Writeln("Document before tracking changes.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert an image while tracking is enabled.
        // The image is a 1x1 pixel PNG encoded in base64 to avoid external files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK0cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imageStream);
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates full lifecycle).
        const string outputPath = "TrackChangesImage.docx";
        doc.Save(outputPath);

        // Verify that the image insertion was recorded as a revision.
        bool insertionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode.NodeType == NodeType.Shape)
            {
                // The inserted shape (image) should also report IsInsertRevision = true.
                Shape insertedShape = (Shape)rev.ParentNode;
                if (insertedShape.IsInsertRevision)
                {
                    insertionFound = true;
                    break;
                }
            }
        }

        if (!insertionFound)
        {
            throw new InvalidOperationException("The image insertion was not recorded as a revision.");
        }

        // Program ends without waiting for user input.
    }
}
