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

        // Insert an inline rectangle shape (width: 100 points, height: 50 points).
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Set the rotation angle to 45 degrees (clockwise).
        shape.Rotation = 45;

        // Define the output file name in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RotatedShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validation: ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Validation: ensure the shape's rotation property is set correctly.
        // Retrieve the shape from the saved document to confirm persistence.
        Document loadedDoc = new Document(outputPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape == null)
            throw new InvalidOperationException("No shape found in the saved document.");

        if (Math.Abs(loadedShape.Rotation - 45) > 0.001)
            throw new InvalidOperationException($"Shape rotation is {loadedShape.Rotation}, expected 45 degrees.");

        // Indicate successful execution (optional, not required for output).
        Console.WriteLine("Document saved and shape rotation verified.");
    }
}
