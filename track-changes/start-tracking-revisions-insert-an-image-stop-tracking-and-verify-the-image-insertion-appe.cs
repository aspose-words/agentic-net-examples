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

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Insert a simple 1x1 PNG image from a base64 string.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imgStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional for verification).
        doc.Save("TrackedChanges.docx");

        // Verify that the image insertion appears as a revision.
        bool imageRevisionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode is Shape)
            {
                imageRevisionFound = true;
                break;
            }
        }

        if (!imageRevisionFound)
            throw new Exception("Image insertion revision was not found.");

        // Indicate successful verification.
        Console.WriteLine("Image insertion revision verified successfully.");
    }
}
