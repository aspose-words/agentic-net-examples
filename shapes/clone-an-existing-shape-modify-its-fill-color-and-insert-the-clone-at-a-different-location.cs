using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class CloneShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an original floating rectangle shape.
        Shape originalShape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
        originalShape.FillColor = Color.Red;
        originalShape.WrapType = WrapType.None;
        originalShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        originalShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        originalShape.Left = 100;
        originalShape.Top = 100;

        // Clone the original shape (deep clone).
        Shape clonedShape = (Shape)originalShape.Clone(true);
        // Change the fill color of the cloned shape.
        clonedShape.FillColor = Color.Blue;
        // Position the cloned shape at a different location.
        clonedShape.Left = 300;
        clonedShape.Top = 300;

        // Insert the cloned shape into the document.
        // Ensure it is attached to the document tree.
        doc.FirstSection.Body.FirstParagraph.AppendChild(clonedShape);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ClonedShape.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
