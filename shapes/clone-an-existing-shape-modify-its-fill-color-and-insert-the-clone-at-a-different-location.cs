using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class CloneShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an original floating rectangle shape.
        Shape originalShape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // 100 points from the left of the page
            RelativeVerticalPosition.Page, 100,     // 100 points from the top of the page
            100, 50,                                // width = 100 points, height = 50 points
            WrapType.None);                         // floating shape, no text wrapping

        // Set the fill color of the original shape.
        originalShape.FillColor = Color.LightBlue;

        // Clone the original shape (deep clone).
        Shape clonedShape = (Shape)originalShape.Clone(true);

        // Change the fill color of the cloned shape.
        clonedShape.FillColor = Color.Yellow;

        // Move the cloned shape to a different location.
        clonedShape.Left = originalShape.Left + 150; // shift right by 150 points
        clonedShape.Top = originalShape.Top;        // keep the same vertical position

        // Insert the cloned shape into the document after the original shape.
        builder.InsertNode(clonedShape);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CloneShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");

        // Optionally, inform that the process completed (no interactive prompts required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
