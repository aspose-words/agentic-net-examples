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
        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.Stroke.Color = System.Drawing.Color.DarkBlue;

        // Insert a floating ellipse shape.
        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 350,
            RelativeVerticalPosition.Page, 150,
            150,
            100,
            WrapType.None);
        ellipse.FillColor = System.Drawing.Color.LightCoral;
        ellipse.Stroke.Color = System.Drawing.Color.Maroon;

        // Insert a line shape.
        Shape line = builder.InsertShape(
            ShapeType.Line,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 300,
            300,   // width (length of the line)
            0,     // height (line has no height)
            WrapType.None);
        line.Stroke.Color = System.Drawing.Color.Green;
        line.StrokeWeight = 2.0;

        // Insert a text box shape and add some text.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 80);
        textBox.FillColor = System.Drawing.Color.LightYellow;
        textBox.Stroke.Color = System.Drawing.Color.Orange;
        // Add a paragraph inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Sample Text Box");
        para.AppendChild(run);

        // Prepare PDF save options to render DrawingML shapes directly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };

        // Define output PDF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapesDocument.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the PDF file.");

        // Optionally, inform that the process completed.
        Console.WriteLine("PDF saved successfully to: " + outputPath);
    }
}
