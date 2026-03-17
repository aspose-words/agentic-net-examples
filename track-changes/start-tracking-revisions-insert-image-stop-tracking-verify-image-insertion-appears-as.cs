using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class RevisionImageExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text (not a revision).
        builder.Writeln("Paragraph before revision tracking.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe");

        // Insert an image while tracking is enabled.
        // The image will be inserted as an inline shape.
        // Replace the path with a valid image file on your system.
        string imagePath = @"C:\Images\SampleImage.jpg";
        builder.InsertImage(imagePath);

        // Stop tracking revisions so subsequent edits are not recorded.
        doc.StopTrackRevisions();

        // Write more text after tracking has stopped.
        builder.Writeln("Paragraph after revision tracking.");

        // Retrieve the inserted shape (image) from the document.
        Shape insertedShape = doc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .FirstOrDefault();

        // Verify that the shape was inserted as a revision.
        if (insertedShape != null && insertedShape.IsInsertRevision)
        {
            Console.WriteLine("The image insertion is recorded as an insert revision.");
        }
        else
        {
            Console.WriteLine("The image insertion is NOT recorded as an insert revision.");
        }

        // Save the document to inspect the result in Microsoft Word if desired.
        doc.Save("RevisionImage.docx");
    }
}
