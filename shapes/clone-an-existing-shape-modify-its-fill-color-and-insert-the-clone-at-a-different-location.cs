using System;
using System.Drawing;
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

        // Insert a floating rectangle shape at a specific position.
        Shape originalShape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // 100 points from the left of the page
            RelativeVerticalPosition.Page, 100,     // 100 points from the top of the page
            150, 100,                               // width = 150 points, height = 100 points
            WrapType.None);                         // No text wrapping

        // Set the fill color of the original shape.
        originalShape.FillColor = Color.LightBlue;

        // Clone the original shape (deep clone).
        Shape clonedShape = (Shape)originalShape.Clone(true);

        // Change the fill color of the cloned shape.
        clonedShape.FillColor = Color.LightCoral;

        // Move the cloned shape to a different location.
        clonedShape.Left = 300; // 300 points from the left of the page
        clonedShape.Top = 200;  // 200 points from the top of the page

        // Insert the cloned shape into the document.
        builder.InsertNode(clonedShape);

        // Define the output file name.
        string outputPath = "ClonedShape.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"The output file '{outputPath}' was not created.");
        }
    }
}
