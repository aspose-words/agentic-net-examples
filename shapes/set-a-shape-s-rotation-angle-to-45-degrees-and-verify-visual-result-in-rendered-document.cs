using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class RotatedShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Set the rotation angle to 45 degrees.
        shape.Rotation = 45f;

        // Verify that the rotation property was set correctly.
        if (Math.Abs(shape.Rotation - 45f) > 0.001f)
            throw new Exception("Shape rotation was not set to 45 degrees.");

        // Save the document to DOCX.
        string docxPath = "RotatedShape.docx";
        doc.Save(docxPath);

        // Save the document to PDF for visual verification.
        string pdfPath = "RotatedShape.pdf";
        doc.Save(pdfPath);

        // Validate that the output files exist.
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("DOCX file was not created.", docxPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);
    }
}
