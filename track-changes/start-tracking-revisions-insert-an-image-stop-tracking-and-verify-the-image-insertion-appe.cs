using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ImageRevisionDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Paragraph before the image.");

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Insert a simple 1x1 PNG image from a byte array.
        // This avoids the need for System.Drawing types.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        builder.InsertImage(imageBytes); // Insertion is recorded as a revision.

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Add more text after the image.
        builder.Writeln("Paragraph after the image.");

        // Verify that the image insertion appears as a revision.
        bool imageRevisionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            // Revisions for inserted shapes (including images) have a Shape as the parent node.
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode is Shape shape && shape.IsInsertRevision)
            {
                imageRevisionFound = true;
                break;
            }
        }

        if (!imageRevisionFound)
            throw new Exception("The inserted image was not recorded as a revision.");

        // Save the document to disk.
        const string outputPath = "ImageRevision.docx";
        doc.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Image insertion revision verified. Document saved to '{outputPath}'.");
    }
}
