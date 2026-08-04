using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape rect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // 100 points from the left of the page
            RelativeVerticalPosition.Page, 100,     // 100 points from the top of the page
            200,                                     // width
            100,                                     // height
            WrapType.None);                         // no text wrapping

        // Set visual properties of the rectangle.
        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 2.0;

        // Insert an inline text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 50);
        textBox.FillColor = System.Drawing.Color.LightYellow;
        textBox.StrokeColor = System.Drawing.Color.Orange;
        textBox.StrokeWeight = 1.5;

        // Add text to the text box.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Size = 12;
        builder.Font.Name = "Arial";
        builder.Writeln("Hello Shapes!");

        // Prepare PDF save options to render DrawingML shapes directly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapesOutput.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // Optionally, inform that the process completed (no interactive I/O required).
        Console.WriteLine("PDF saved successfully to: " + outputPath);
    }
}
