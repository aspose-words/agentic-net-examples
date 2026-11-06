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

        // Start tracking revisions with a specific author and date.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Use DocumentBuilder to insert an image while tracking is enabled.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is an inserted image:");

        // Small 1x1 pixel PNG image (base64 encoded).
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK0cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            // Insert the image; this operation will be recorded as a revision.
            builder.InsertImage(imageStream);
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document to verify the output file is created.
        string outputPath = "TrackChangesImage.docx";
        doc.Save(outputPath);

        // Verify that the image insertion appears as a revision.
        bool imageInsertionRevisionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode is Shape)
            {
                imageInsertionRevisionFound = true;
                break;
            }
        }

        if (!imageInsertionRevisionFound)
        {
            throw new InvalidOperationException("The image insertion revision was not found.");
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Image insertion revision verified successfully.");
    }
}
