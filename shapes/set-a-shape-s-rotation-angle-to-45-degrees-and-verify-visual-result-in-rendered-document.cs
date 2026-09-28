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

        // Insert a rectangle shape (inline by default) with specific size.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 100);

        // Set the rotation angle to 45 degrees.
        shape.Rotation = 45;

        // Validate that the rotation was applied.
        if (Math.Abs(shape.Rotation - 45) > 0.001)
            throw new Exception("Shape rotation was not set to 45 degrees.");

        // Save the document to a .docx file.
        string docPath = "RotatedShape.docx";
        doc.Save(docPath);

        // Render the first page of the document to a PNG image to visually verify the rotation.
        string imagePath = "RotatedShape.png";
        doc.Save(imagePath, SaveFormat.Png);

        // Verify that both output files exist.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("The output DOCX file was not created.", docPath);
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("The output PNG file was not created.", imagePath);
    }
}
