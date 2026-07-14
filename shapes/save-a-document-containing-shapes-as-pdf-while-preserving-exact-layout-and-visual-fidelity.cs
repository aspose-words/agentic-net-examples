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
            RelativeVerticalPosition.Page, 100,    // 100 points from the top of the page
            200,                                    // width
            100,                                    // height
            WrapType.None);                         // no text wrapping

        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 2.0;

        // Insert a text box shape (inline) and add some text to it.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 80);
        textBox.FillColor = System.Drawing.Color.LightYellow;
        textBox.StrokeColor = System.Drawing.Color.Orange;
        textBox.StrokeWeight = 1.5;

        // Add a paragraph inside the text box.
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Hello Aspose.Words!");
        run.Font.Size = 14;
        run.Font.Bold = true;
        para.AppendChild(run);

        // Create a simple in‑memory PNG image (a single green pixel) using a base‑64 string.
        // This avoids the need for System.Drawing.Bitmap which may not be available on all platforms.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8z8BQDwAFgwJ/lKXK5wAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the PNG as an inline image.
        Shape imageShape = builder.InsertImage(imageBytes);
        imageShape.WrapType = WrapType.None;
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        imageShape.Left = 350; // position it to the right of the rectangle
        imageShape.Top = 120;

        // Prepare PDF save options to render DrawingML shapes directly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapesDocument.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF file.");
    }
}
