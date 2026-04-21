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

        // Insert a rectangle shape with fill and stroke.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        rectangle.FillColor = System.Drawing.Color.LightBlue;
        rectangle.StrokeColor = System.Drawing.Color.DarkBlue;
        rectangle.StrokeWeight = 2.0;

        // Insert a text box shape with some text.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 250;
        textBox.Height = 80;
        textBox.WrapType = WrapType.None;
        textBox.FillColor = System.Drawing.Color.LightYellow;
        textBox.StrokeColor = System.Drawing.Color.Orange;
        textBox.StrokeWeight = 1.5;

        // Add a paragraph and run inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello Aspose.Words Shapes!");
        run.Font.Size = 14;
        run.Font.Bold = true;
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Append the text box to the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

        // Create a simple 1x1 red PNG image (base64 encoded) in memory.
        // This avoids the need for System.Drawing dependencies.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK5cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the image as a floating shape centered on the page.
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            Shape imageShape = builder.InsertImage(imgStream);
            imageShape.WrapType = WrapType.None;
            imageShape.BehindText = true;
            imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            imageShape.HorizontalAlignment = HorizontalAlignment.Center;
            imageShape.VerticalAlignment = VerticalAlignment.Center;
        }

        // Define the output PDF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapesOutput.pdf");

        // Save the document as PDF with DrawingML rendering to preserve visual fidelity.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the PDF file.");

        // Indicate success.
        Console.WriteLine("Document with shapes saved as PDF: " + outputPath);
    }
}
