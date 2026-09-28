using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an original rectangle shape.
        Shape originalShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        originalShape.FillColor = Color.Blue;

        // Clone the original shape (deep clone).
        Shape clonedShape = (Shape)originalShape.Clone(true);

        // Modify the clone's fill color.
        clonedShape.FillColor = Color.Green;

        // Configure the clone as a floating shape and position it elsewhere.
        clonedShape.WrapType = WrapType.None;
        clonedShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        clonedShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        clonedShape.Left = 150;   // Points from the left edge of the page.
        clonedShape.Top = 150;    // Points from the top edge of the page.

        // Insert the cloned shape at the end of the document.
        builder.MoveToDocumentEnd();
        builder.InsertNode(clonedShape);

        // Save the document.
        string outputPath = "ClonedShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");
    }
}
