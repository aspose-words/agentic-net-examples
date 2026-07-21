using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an initial floating rectangle shape.
        Shape original = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left position
            RelativeVerticalPosition.Page, 100,     // top position
            100,                                     // width
            50,                                      // height
            WrapType.None);

        // Set the fill color of the original shape.
        original.FillColor = Color.LightGray;

        // Clone the original shape (deep clone).
        Shape cloned = (Shape)original.Clone(true);

        // Modify the fill color of the cloned shape.
        cloned.FillColor = Color.LightBlue;

        // Position the cloned shape at a different location.
        cloned.Left = 300; // points from the left edge of the page
        cloned.Top = 200;  // points from the top edge of the page

        // Insert the cloned shape into the document tree.
        doc.FirstSection.Body.FirstParagraph.AppendChild(cloned);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ClonedShape.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");
    }
}
