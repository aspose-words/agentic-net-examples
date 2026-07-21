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
            200, 100,                               // width = 200 points, height = 100 points
            WrapType.None);                         // No text wrapping

        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 2.0;

        // Insert a floating text box shape with some text.
        Shape textBox = builder.InsertShape(
            ShapeType.TextBox,
            RelativeHorizontalPosition.Page, 350,
            RelativeVerticalPosition.Page, 150,
            250, 100,
            WrapType.None);

        textBox.FillColor = System.Drawing.Color.LightYellow;
        textBox.StrokeColor = System.Drawing.Color.Orange;
        textBox.StrokeWeight = 1.5;

        // Add a paragraph inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello Aspose.Words Shapes!");
        run.Font.Size = 14;
        run.Font.Bold = true;
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Configure PDF save options to render DrawingML shapes directly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };

        // Determine output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Shapes.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // Optionally, you could add further validation of file size, etc.
    }
}
