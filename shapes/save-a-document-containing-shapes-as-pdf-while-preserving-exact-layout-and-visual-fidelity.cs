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

        // Insert a paragraph to hold the shapes.
        builder.Writeln("Document with shapes saved as PDF:");
        builder.Writeln();

        // Insert a floating rectangle shape.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 150;
        rectangle.Height = 100;
        rectangle.Left = 100; // points from the left edge of the page
        rectangle.Top = 100;  // points from the top edge of the page
        rectangle.WrapType = WrapType.None; // No text wrapping
        rectangle.StrokeColor = System.Drawing.Color.Blue;
        rectangle.FillColor = System.Drawing.Color.LightBlue;
        rectangle.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        rectangle.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Append the rectangle to the document body.
        builder.CurrentParagraph.AppendChild(rectangle);
        builder.Writeln(); // Add space after the shape.

        // Insert a floating ellipse shape.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 120;
        ellipse.Height = 80;
        ellipse.Left = 300;
        ellipse.Top = 150;
        ellipse.WrapType = WrapType.None;
        ellipse.StrokeColor = System.Drawing.Color.Green;
        ellipse.FillColor = System.Drawing.Color.LightGreen;
        ellipse.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        ellipse.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Append the ellipse to the document body.
        builder.CurrentParagraph.AppendChild(ellipse);
        builder.Writeln(); // Add space after the shape.

        // Define output paths.
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeShapesExample");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "ShapesDocument.pdf");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");
        }

        // Optionally, clean up (comment out if you want to inspect the file).
        // File.Delete(pdfPath);
    }
}
